#!/usr/bin/env python
# DaVinci Resolve - Align Timeline Contents to AAF by Timecode
# =============================================================
# Reads audio clips from an AAF timeline, finds matching source timelines
# in a bin (by timecode), copies the corresponding video clips to the
# destination timeline at the correct record positions, copies markers,
# and restores multicam active angles via DRT export/import.
#
# Requirements:
#   - Open the AAF timeline (audio-only) as the current timeline
#   - Place source timelines in a bin named "TRACKS" (or select via dialog)
#   - TC offset is derived from clip media TC

import copy
import os
import re
import zipfile
import tempfile

DEFAULT_CONFIG = {
    'video_tracks_count': 1,
    'ignore_prefixes': ["Sample", "Fade"],
    'create_new_timeline': True,
    'new_timeline_suffix': '_montaz',
}

_config = copy.deepcopy(DEFAULT_CONFIG)


def get_resolve():
    try:
        import DaVinciResolveScript as dvr
        return dvr.scriptapp("Resolve")
    except ImportError:
        try:
            return bmd.scriptapp("Resolve")
        except NameError:
            return None


def get_fusion():
    try:
        return bmd.scriptapp("Fusion")
    except (NameError, AttributeError):
        return None


def tc_to_frames(timecode, fps):
    if not fps or fps <= 0:
        return None
    parts = timecode.replace(';', ':').split(':')
    if len(parts) != 4:
        return None
    try:
        h, m, s, f = map(int, parts)
    except ValueError:
        return None
    fps_int = int(round(fps))
    return ((h * 3600 + m * 60 + s) * fps_int) + f


def frames_to_tc(frames, fps):
    if not fps or fps <= 0:
        return "00:00:00:00"
    fps_int = int(round(fps))
    f = frames % fps_int
    s = (frames // fps_int) % 60
    m = (frames // (fps_int * 60)) % 60
    h = frames // (fps_int * 3600)
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(h, m, s, f)


def should_skip_clip(name):
    for prefix in _config.get('ignore_prefixes', []):
        if name.startswith(prefix):
            return True
    return False


def get_all_bins(folder, bins_list, path=""):
    current_path = (path + "/" + folder.GetName()) if path else folder.GetName()
    clips = folder.GetClipList()
    if clips:
        bins_list.append({
            'folder': folder,
            'name': folder.GetName(),
            'path': current_path,
            'clip_count': len(clips),
        })
    for sub in folder.GetSubFolderList():
        get_all_bins(sub, bins_list, current_path)


def load_source_timelines(bin_folder, project, fps):
    """Match timeline items in bin to Timeline objects from the project."""
    clips = bin_folder.GetClipList()
    if not clips:
        return []

    tl_names = set()
    for item in clips:
        if item.GetClipProperty("Type") == "Timeline":
            tl_names.add(item.GetName())

    if not tl_names:
        for item in clips:
            tl_names.add(item.GetName())

    result = []
    count = project.GetTimelineCount()
    for i in range(1, count + 1):
        tl = project.GetTimelineByIndex(i)
        if not tl or tl.GetName() not in tl_names:
            continue

        # GetStartTimecode() gives the absolute TC of the timeline's first frame.
        start_tc_str = tl.GetStartTimecode() or "00:00:00:00"
        tc_start = tc_to_frames(start_tc_str, fps)
        if tc_start is None:
            tc_start = 0

        duration = tl.GetEndFrame() - tl.GetStartFrame()
        tc_end = tc_start + duration

        result.append({
            'timeline': tl,
            'name': tl.GetName(),
            'tc_start': tc_start,
            'tc_end': tc_end,
        })
        print("  {} | {} - {}".format(
            tl.GetName(),
            frames_to_tc(tc_start, fps),
            frames_to_tc(tc_end, fps),
        ))

    return result


def find_source_timeline(tc_frames, source_timelines):
    for src in source_timelines:
        if src['tc_start'] <= tc_frames < src['tc_end']:
            return src
    return None


# ---------------------------------------------------------------------------
# Multicam angle restoration via DRT
# ---------------------------------------------------------------------------
# DRT files are ZIP archives containing XML. The active multicam angle for
# each placed clip is stored in a protobuf FieldsBlob inside the clip's
# color version entry. The camera name "Camera X" (X=1-9) is encoded as
# ASCII within that blob:  1a=field-tag  08=length-8  43616d657261 20=b"Camera "  3X=digit
# We read angles from source DRTs, track where each clip lands in the
# destination, then export the destination DRT, patch the blobs, and
# reimport — bypassing the missing scripting API for multicam angles.

def _parse_drt_camera_angles(drt_path):
    """
    Open a DRT file and return {start_frame: camera_num} for every video
    clip whose active angle was found in a FieldsBlob.
    """
    result = {}
    try:
        with zipfile.ZipFile(drt_path, 'r') as z:
            seq_files = [
                (name, z.getinfo(name).file_size)
                for name in z.namelist()
                if name.startswith('SeqContainer/') and name.endswith('.xml')
            ]
            if not seq_files:
                return result
            main_xml = max(seq_files, key=lambda x: x[1])[0]
            xml_content = z.read(main_xml).decode('utf-8')
    except Exception as e:
        print("    DRT read error: {}".format(e))
        return result

    for clip_m in re.finditer(
            r'<Sm2TiVideoClip[^>]*>(.*?)</Sm2TiVideoClip>',
            xml_content, re.DOTALL):
        block = clip_m.group(1)
        start_m = re.search(r'<Start>(\d+)</Start>', block)
        if not start_m:
            continue
        start_frame = int(start_m.group(1))

        # Active angle lives in LmVersion; fall back to first matching blob
        lm_m = re.search(
            r'<ListMgt::LmVersion[^>]*>(.*?)</ListMgt::LmVersion>',
            block, re.DOTALL)
        search_in = lm_m.group(1) if lm_m else block

        for blob_m in re.finditer(r'<FieldsBlob>([0-9a-fA-F]+)</FieldsBlob>', search_in):
            blob = blob_m.group(1).lower()
            cam_m = re.search(r'1a0843616d65726120(3[1-9])', blob)
            if cam_m:
                result[start_frame] = int(cam_m.group(1), 16) - 0x30
                break

    return result


def _set_camera_in_blob(blob_hex, camera_num):
    """Replace the camera angle digit in a FieldsBlob hex string (1-9)."""
    if camera_num < 1 or camera_num > 9:
        return blob_hex
    new_digit = '{:02x}'.format(0x30 + camera_num)
    return re.sub(
        r'(1a0843616d65726120)3[0-9]',
        lambda m: m.group(1) + new_digit,
        blob_hex, flags=re.IGNORECASE, count=1,
    )


def build_source_camera_map(source_timelines, resolve, temp_dir):
    """
    Export each source timeline as DRT and parse camera angle data.
    Returns {tl_name: {start_frame: camera_num}}.
    """
    angle_map = {}
    for src in source_timelines:
        tl = src['timeline']
        tl_name = src['name']
        safe_name = re.sub(r'[^\w-]', '_', tl_name)
        drt_path = os.path.join(temp_dir, safe_name + '_src.drt')
        try:
            ok = tl.Export(drt_path, resolve.EXPORT_DRT, resolve.EXPORT_NONE)
            if not ok:
                print("  WARNING: DRT export failed for '{}'".format(tl_name))
                continue
            clips_map = _parse_drt_camera_angles(drt_path)
            if clips_map:
                angle_map[tl_name] = clips_map
                print("  DRT angles: '{}' -> {} clip(s)".format(
                    tl_name, len(clips_map)))
        except Exception as e:
            print("  WARNING: DRT angle parse error for '{}': {}".format(tl_name, e))
    return angle_map


def fix_multicam_angles_via_drt(dest_timeline, dest_camera_map, project,
                                  media_pool, resolve, temp_dir):
    """
    Patch multicam angles by exporting the destination timeline as DRT,
    modifying FieldsBlob camera values, deleting the original timeline,
    and reimporting the patched DRT.
    Returns the new Timeline object, or the original on error.
    """
    to_fix = {k: v for k, v in dest_camera_map.items() if v != 1}
    if not to_fix:
        return dest_timeline

    tl_name = dest_timeline.GetName()
    safe = re.sub(r'[^\w-]', '_', tl_name)
    drt_orig  = os.path.join(temp_dir, safe + '_dest.drt')
    drt_fixed = os.path.join(temp_dir, safe + '_dest_fixed.drt')

    print("\nApplying multicam angles via DRT ({} clips to fix)...".format(len(to_fix)))

    if not dest_timeline.Export(drt_orig, resolve.EXPORT_DRT, resolve.EXPORT_NONE):
        print("  ERROR: DRT export of destination failed - angles not applied")
        return dest_timeline

    with zipfile.ZipFile(drt_orig, 'r') as z:
        file_data = {name: z.read(name) for name in z.namelist()}

    fixed_count = [0]

    def patch_clip(m):
        block = m.group(0)
        start_m = re.search(r'<Start>(\d+)</Start>', block)
        if not start_m:
            return block
        cam = to_fix.get(int(start_m.group(1)))
        if cam is None:
            return block

        def patch_blob(bm):
            new_hex = _set_camera_in_blob(bm.group(1).lower(), cam)
            if new_hex != bm.group(1).lower():
                fixed_count[0] += 1
                return '<FieldsBlob>' + new_hex + '</FieldsBlob>'
            return bm.group(0)

        return re.sub(r'<FieldsBlob>([0-9a-fA-F]*)</FieldsBlob>', patch_blob, block)

    for fname in list(file_data.keys()):
        if fname.startswith('SeqContainer/') and fname.endswith('.xml'):
            xml_str = file_data[fname].decode('utf-8')
            patched = re.sub(
                r'<Sm2TiVideoClip[^>]*>.*?</Sm2TiVideoClip>',
                patch_clip, xml_str, flags=re.DOTALL,
            )
            file_data[fname] = patched.encode('utf-8')

    with zipfile.ZipFile(drt_fixed, 'w', zipfile.ZIP_DEFLATED) as z:
        for fname, content in file_data.items():
            z.writestr(fname, content)

    print("  Patched {} FieldsBlob(s)".format(fixed_count[0]))

    if not media_pool.DeleteTimelines([dest_timeline]):
        print("  ERROR: Could not delete destination timeline - angles not applied")
        return dest_timeline

    new_tl = media_pool.ImportTimelineFromFile(drt_fixed, {})
    if not new_tl:
        print("  CRITICAL: DRT reimport failed - original timeline was deleted!")
        return None

    if new_tl.GetName() != tl_name:
        new_tl.SetName(tl_name)

    project.SetCurrentTimeline(new_tl)
    print("  Multicam angles applied -> '{}'".format(new_tl.GetName()))
    return new_tl


# ---------------------------------------------------------------------------
# Core clip / marker / timeline operations
# ---------------------------------------------------------------------------

def copy_clips_from_source(src_data, tc_in, tc_out, record_start,
                            media_pool, fps, num_tracks, source_angle_map=None):
    """
    Copy video clips from src_data timeline where source TC overlaps [tc_in, tc_out).
    Returns (placed, failed, angle_assignments).
    angle_assignments: {dest_start_frame: camera_num} for clips that need a fix.
    """
    tl = src_data['timeline']
    placed = 0
    failed = 0
    angle_assignments = {}

    for track_idx in range(1, num_tracks + 1):
        items = tl.GetItemListInTrack("video", track_idx)
        if not items:
            if track_idx > 1:
                print("    V{}: no clips in source timeline".format(track_idx))
            continue

        track_placed = 0
        for item in items:
            media_item = item.GetMediaPoolItem()
            if not media_item:
                continue

            # item.GetStart() / GetEnd() return absolute TC frame numbers
            item_tc_in  = item.GetStart()
            item_tc_out = item.GetEnd()

            overlap_in  = max(item_tc_in, tc_in)
            overlap_out = min(item_tc_out, tc_out)
            if overlap_in >= overlap_out:
                continue

            left_offset  = item.GetLeftOffset()
            src_in       = left_offset + (overlap_in - item_tc_in)
            src_out      = src_in + (overlap_out - overlap_in)
            record_frame = record_start + (overlap_in - tc_in)

            clip_info = {
                "mediaPoolItem": media_item,
                "startFrame":    src_in,
                "endFrame":      src_out,
                "mediaType":     1,
                "trackIndex":    track_idx,
                "recordFrame":   record_frame,
            }

            result = media_pool.AppendToTimeline([clip_info])
            if result:
                placed += 1
                track_placed += 1
                if source_angle_map is not None:
                    cam = source_angle_map.get(item_tc_in, 1)
                    if cam != 1:
                        angle_assignments[result[0].GetStart()] = cam
            else:
                failed += 1
                print("    FAILED: {} on V{}".format(item.GetName(), track_idx))

        if track_idx > 1 and track_placed > 0:
            print("    V{}: placed {}".format(track_idx, track_placed))

    return placed, failed, angle_assignments


def copy_markers_from_source(src_data, tc_in, tc_out, record_start, dest_timeline, fps):
    """
    Copy markers from src_data timeline that fall within [tc_in, tc_out)
    to dest_timeline, preserving their relative position.
    Returns number of markers copied.
    """
    tl = src_data['timeline']
    src_tc_start = src_data['tc_start']

    markers = tl.GetMarkers()
    if not markers:
        return 0

    copied = 0
    for frame_pos, data in markers.items():
        # GetMarkers() frame_pos is relative to timeline start (0-based),
        # unlike item.GetStart() which returns absolute TC frames.
        marker_tc = src_tc_start + frame_pos
        if not (tc_in <= marker_tc < tc_out):
            continue

        dest_frame = record_start + (marker_tc - tc_in)

        max_duration = tc_out - marker_tc
        duration = min(data.get('duration', 1), max_duration)
        duration = max(duration, 1)

        success = dest_timeline.AddMarker(
            dest_frame,
            data.get('color', 'Blue'),
            data.get('name', ''),
            data.get('note', ''),
            duration,
            data.get('customData', ''),
        )
        if success:
            copied += 1

    return copied


def create_dest_timeline(project, media_pool, aaf_timeline, suffix):
    new_name = aaf_timeline.GetName() + suffix
    fps      = aaf_timeline.GetSetting("timelineFrameRate")
    start_tc = aaf_timeline.GetStartTimecode()

    new_tl = media_pool.CreateEmptyTimeline(new_name)
    if not new_tl:
        return None

    new_tl.SetSetting("timelineFrameRate", fps)
    if start_tc:
        new_tl.SetStartTimecode(start_tc)

    print("Created: '{}' | {} fps | start TC {}".format(new_name, fps, start_tc))
    return new_tl


def copy_audio_from_aaf(aaf_tl, dest_tl, media_pool):
    audio_count = aaf_tl.GetTrackCount("audio")
    dest_count  = dest_tl.GetTrackCount("audio")
    while dest_count < audio_count:
        dest_tl.AddTrack("audio")
        new_count = dest_tl.GetTrackCount("audio")
        if new_count == dest_count:
            break
        dest_count = new_count

    copied = 0
    for track_idx in range(1, audio_count + 1):
        items = aaf_tl.GetItemListInTrack("audio", track_idx)
        if not items:
            continue
        for item in items:
            if should_skip_clip(item.GetName()):
                continue
            mpi = item.GetMediaPoolItem()
            if not mpi:
                continue
            clip_info = {
                "mediaPoolItem": mpi,
                "startFrame":    item.GetLeftOffset(),
                "endFrame":      item.GetLeftOffset() + item.GetDuration(),
                "trackIndex":    track_idx,
                "recordFrame":   item.GetStart(),
            }
            if media_pool.AppendToTimeline([clip_info]):
                copied += 1

    print("Copied {} audio clips from AAF".format(copied))
    return copied


def ensure_video_tracks(timeline, count):
    current = timeline.GetTrackCount("video")
    while current < count:
        timeline.AddTrack("video")
        new_count = timeline.GetTrackCount("video")
        if new_count == current:
            print("WARNING: Could not add more video tracks (stuck at {})".format(current))
            break
        current = new_count


def show_settings_dialog(bins, fusion):
    global _config

    ui   = fusion.UIManager
    disp = bmd.UIDispatcher(ui)
    result = {'bin_idx': -1, 'cancelled': True}

    win = disp.AddWindow({
        'ID': 'AlignWin',
        'WindowTitle': 'Align Timelines to AAF - Settings',
        'Geometry': [300, 200, 480, 400],
        'Spacing': 10,
    }, [
        ui.VGroup({'Spacing': 6}, [
            ui.Label({
                'Text': 'Align Timeline Contents to AAF',
                'Font': ui.Font({'PixelSize': 15, 'Bold': True}),
                'Weight': 0,
            }),
            ui.Label({'Text': '-' * 55, 'Weight': 0}),

            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Source timelines bin:', 'Weight': 2}),
                ui.ComboBox({'ID': 'BinCombo', 'Weight': 3}),
            ]),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Copy video tracks V1 through:', 'Weight': 2}),
                ui.SpinBox({
                    'ID': 'TrackCount',
                    'Value': _config.get('video_tracks_count', 1),
                    'Minimum': 1,
                    'Maximum': 20,
                    'Weight': 1,
                }),
            ]),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Ignore clip prefixes (comma-sep):', 'Weight': 2}),
                ui.LineEdit({
                    'ID': 'IgnorePrefixes',
                    'Text': ', '.join(_config.get('ignore_prefixes', [])),
                    'Weight': 2,
                }),
            ]),

            ui.Label({'Text': '-' * 55, 'Weight': 0}),
            ui.Label({'Text': 'Destination Timeline:', 'Font': ui.Font({'Bold': True}), 'Weight': 0}),
            ui.HGroup({'Weight': 0}, [
                ui.CheckBox({
                    'ID': 'CreateNew',
                    'Text': 'Create new timeline (recommended)',
                    'Checked': _config.get('create_new_timeline', True),
                    'Weight': 1,
                }),
            ]),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'New timeline name suffix:', 'Weight': 2}),
                ui.LineEdit({
                    'ID': 'Suffix',
                    'Text': _config.get('new_timeline_suffix', '_montaz'),
                    'Weight': 2,
                }),
            ]),

            ui.Label({'Text': '', 'Weight': 1}),
            ui.HGroup({'Weight': 0}, [
                ui.Button({'ID': 'StartBtn', 'Text': 'Start', 'Weight': 1}),
                ui.Button({'ID': 'CancelBtn', 'Text': 'Cancel', 'Weight': 1}),
            ]),
        ]),
    ])

    combo = win.Find('BinCombo')
    default_idx = 0
    for i, b in enumerate(bins):
        combo.AddItem("{} ({} clips)".format(b['name'], b['clip_count']))
        if b['name'].upper() == 'TRACKS':
            default_idx = i
    if bins:
        combo.CurrentIndex = default_idx

    def on_start(ev):
        _config['video_tracks_count']  = win.Find('TrackCount').Value
        raw = win.Find('IgnorePrefixes').Text
        _config['ignore_prefixes']     = [p.strip() for p in raw.split(',') if p.strip()]
        _config['create_new_timeline'] = win.Find('CreateNew').Checked
        _config['new_timeline_suffix'] = win.Find('Suffix').Text
        result['bin_idx']   = win.Find('BinCombo').CurrentIndex
        result['cancelled'] = False
        disp.ExitLoop()

    def on_cancel(ev):
        disp.ExitLoop()

    def on_close(ev):
        disp.ExitLoop()

    win.On.StartBtn.Clicked  = on_start
    win.On.CancelBtn.Clicked = on_cancel
    win.On.AlignWin.Close    = on_close

    win.Show()
    disp.RunLoop()
    win.Hide()

    return -1 if result['cancelled'] else result['bin_idx']


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("")
    print("=" * 60)
    print("  Align Timeline Contents to AAF by Timecode")
    print("=" * 60)

    resolve = get_resolve()
    if not resolve:
        print("ERROR: Could not connect to DaVinci Resolve")
        return

    project = resolve.GetProjectManager().GetCurrentProject()
    if not project:
        print("ERROR: No project open")
        return

    media_pool   = project.GetMediaPool()
    aaf_timeline = project.GetCurrentTimeline()
    if not aaf_timeline:
        print("ERROR: No timeline open. Open the AAF timeline first.")
        return

    fps = float(aaf_timeline.GetSetting("timelineFrameRate"))
    print("")
    print("AAF timeline : {}".format(aaf_timeline.GetName()))
    print("Frame rate   : {} fps".format(fps))

    all_bins = []
    get_all_bins(media_pool.GetRootFolder(), all_bins)
    if not all_bins:
        print("ERROR: No bins with clips found in media pool")
        return

    fusion       = get_fusion()
    selected_idx = -1

    if fusion:
        try:
            selected_idx = show_settings_dialog(all_bins, fusion)
        except Exception as e:
            print("Settings dialog failed: {}".format(e))

    if selected_idx < 0:
        for i, b in enumerate(all_bins):
            if b['name'].upper() == 'TRACKS':
                selected_idx = i
                print("Auto-selected bin: '{}'".format(b['name']))
                break

    if selected_idx < 0 or selected_idx >= len(all_bins):
        print("ERROR: No bin selected.")
        print("       Rename your source timelines bin to 'TRACKS' for auto-detection.")
        return

    source_bin = all_bins[selected_idx]['folder']
    num_tracks = _config['video_tracks_count']
    create_new = _config['create_new_timeline']
    suffix     = _config['new_timeline_suffix']

    print("")
    print("Source bin   : {}".format(source_bin.GetName()))
    print("Video tracks : V1 - V{}".format(num_tracks))
    print("")
    print("Loading source timelines...")

    source_timelines = load_source_timelines(source_bin, project, fps)
    if not source_timelines:
        print("ERROR: No valid timelines found in bin.")
        return

    # Create temp dir for DRT operations
    temp_dir = tempfile.mkdtemp(prefix='resolve_align_')

    # Parse camera angles from source DRTs before placing any clips
    print("")
    print("Loading multicam angle data from source DRTs...")
    source_camera_map = build_source_camera_map(source_timelines, resolve, temp_dir)
    if not source_camera_map:
        print("  (no multicam angle data found - skipping angle fix)")

    print("")
    print("Reading AAF audio track 1...")
    audio_items = aaf_timeline.GetItemListInTrack("audio", 1)
    if not audio_items:
        print("ERROR: No clips on audio track 1 of the AAF timeline")
        return
    print("Found {} audio clips".format(len(audio_items)))

    dest_timeline = aaf_timeline
    if create_new:
        print("")
        dest_timeline = create_dest_timeline(project, media_pool, aaf_timeline, suffix)
        if not dest_timeline:
            print("ERROR: Failed to create destination timeline")
            return
        project.SetCurrentTimeline(dest_timeline)
        print("")
        copy_audio_from_aaf(aaf_timeline, dest_timeline, media_pool)

    ensure_video_tracks(dest_timeline, num_tracks)

    print("")
    print("Aligning...")
    print("-" * 60)

    total_placed   = 0
    total_failed   = 0
    total_markers  = 0
    no_match_count = 0
    dest_camera_map = {}  # {dest_start_frame: camera_num}

    for audio_item in audio_items:
        clip_name = audio_item.GetName()

        if should_skip_clip(clip_name):
            print("SKIP  : {}".format(clip_name))
            continue

        # Read source TC - works even for offline clips via project DB
        tc_str = None
        mpi = audio_item.GetMediaPoolItem()
        if mpi:
            tc_str = mpi.GetClipProperty("Start TC")
        if not tc_str:
            gcp = getattr(audio_item, 'GetClipProperty', None)
            if callable(gcp):
                tc_str = gcp("Start TC")
        if not tc_str:
            print("NO TC : {}".format(clip_name))
            continue

        clip_start_frames = tc_to_frames(tc_str, fps)
        if clip_start_frames is None:
            continue

        left_offset  = audio_item.GetLeftOffset()
        record_start = audio_item.GetStart()
        duration     = audio_item.GetDuration()

        tc_in  = clip_start_frames + left_offset
        tc_out = tc_in + duration

        src = find_source_timeline(tc_in, source_timelines)
        if not src:
            print("NO MATCH: {} (TC {})".format(clip_name, frames_to_tc(tc_in, fps)))
            no_match_count += 1
            continue

        print("MATCH : {} -> '{}' [{}]".format(
            clip_name, src['name'], frames_to_tc(tc_in, fps),
        ))

        src_angle_map = source_camera_map.get(src['name'], {})

        placed, failed, angles = copy_clips_from_source(
            src, tc_in, tc_out, record_start,
            media_pool, fps, num_tracks, src_angle_map,
        )
        total_placed += placed
        total_failed += failed
        dest_camera_map.update(angles)

        markers_copied = copy_markers_from_source(
            src, tc_in, tc_out, record_start, dest_timeline, fps,
        )
        total_markers += markers_copied
        if markers_copied:
            print("  + {} marker(s)".format(markers_copied))

    print("-" * 60)

    # Restore multicam angles via DRT patch
    if dest_camera_map:
        dest_timeline = fix_multicam_angles_via_drt(
            dest_timeline, dest_camera_map,
            project, media_pool, resolve, temp_dir,
        )

    # Clean up temp files
    try:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    print("")
    print("=" * 60)
    print("  DONE")
    print("  Video clips placed : {}".format(total_placed))
    if total_failed:
        print("  Placement failures : {}".format(total_failed))
    print("  Markers copied     : {}".format(total_markers))
    if dest_camera_map:
        non1 = len([v for v in dest_camera_map.values() if v != 1])
        print("  Multicam angles fixed : {}".format(non1))
    if no_match_count:
        print("  Audio clips with no TC match : {}".format(no_match_count))
    if create_new and dest_timeline:
        print("  Destination timeline : '{}'".format(dest_timeline.GetName()))
    print("=" * 60)
    print("")


main()
