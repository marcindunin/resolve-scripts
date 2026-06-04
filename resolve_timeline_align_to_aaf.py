#!/usr/bin/env python
# DaVinci Resolve - Align Timeline Contents to AAF by Timecode
# =============================================================
# Reads audio clips from an AAF timeline, finds matching source timelines
# in a bin (by timecode), copies the corresponding video clips to the
# destination timeline at the correct record positions, and copies markers.
#
# Requirements:
#   - Open the AAF timeline (audio-only) as the current timeline
#   - Place source timelines in a bin named "TRACKS" (or select via dialog)
#   - TC offset is derived from clip media TC, so GetStartTimecode() does not
#     need to be set correctly on the source timelines

import copy

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
        # Fallback: treat all items in bin as potential timelines
        for item in clips:
            tl_names.add(item.GetName())

    result = []
    count = project.GetTimelineCount()
    for i in range(1, count + 1):
        tl = project.GetTimelineByIndex(i)
        if not tl or tl.GetName() not in tl_names:
            continue

        # GetStartTimecode() gives the absolute TC of the timeline's first frame.
        # Clips inside the timeline may have relative TC (starting at 00:00:00:00),
        # so we must use the timeline's Start TC — not clip media TC — for positioning.
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


def copy_clips_from_source(src_data, tc_in, tc_out, record_start,
                            media_pool, fps, num_tracks):
    """
    Copy video clips from src_data timeline where source TC overlaps [tc_in, tc_out).
    Clip TC = timeline Start TC + item.GetStart() — correct for multitrack clips
    whose media has relative TC (starting at 00:00:00:00).
    record_start: destination frame (0-based from dest timeline start) where tc_in lands.
    Returns (placed, failed).
    """
    tl = src_data['timeline']
    src_tc_start = src_data['tc_start']
    placed = 0
    failed = 0

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

            # Absolute TC of this clip = timeline Start TC + its position in the timeline
            item_tc_in = src_tc_start + item.GetStart()
            item_tc_out = src_tc_start + item.GetEnd()

            # Intersection with requested TC range
            overlap_in = max(item_tc_in, tc_in)
            overlap_out = min(item_tc_out, tc_out)
            if overlap_in >= overlap_out:
                continue

            # Source in/out within the media file
            left_offset = item.GetLeftOffset()
            src_in = left_offset + (overlap_in - item_tc_in)
            src_out = src_in + (overlap_out - overlap_in)

            # Record position in destination timeline
            record_frame = record_start + (overlap_in - tc_in)

            clip_info = {
                "mediaPoolItem": media_item,
                "startFrame": src_in,
                "endFrame": src_out,
                "mediaType": 1,
                "trackIndex": track_idx,
                "recordFrame": record_frame,
            }

            if media_pool.AppendToTimeline([clip_info]):
                placed += 1
                track_placed += 1
            else:
                failed += 1
                print("    FAILED: {} on V{}".format(item.GetName(), track_idx))

        if track_idx > 1 and track_placed > 0:
            print("    V{}: placed {}".format(track_idx, track_placed))

    return placed, failed


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
        marker_tc = src_tc_start + frame_pos
        if not (tc_in <= marker_tc < tc_out):
            continue

        dest_frame = record_start + (marker_tc - tc_in)

        # Clamp marker duration so it doesn't exceed the copied range
        max_duration = (tc_out - marker_tc)
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
    fps = aaf_timeline.GetSetting("timelineFrameRate")
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
    dest_count = dest_tl.GetTrackCount("audio")
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
                "startFrame": item.GetLeftOffset(),
                "endFrame": item.GetLeftOffset() + item.GetDuration(),
                "trackIndex": track_idx,
                "recordFrame": item.GetStart(),
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

    ui = fusion.UIManager
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
        _config['video_tracks_count'] = win.Find('TrackCount').Value
        raw = win.Find('IgnorePrefixes').Text
        _config['ignore_prefixes'] = [p.strip() for p in raw.split(',') if p.strip()]
        _config['create_new_timeline'] = win.Find('CreateNew').Checked
        _config['new_timeline_suffix'] = win.Find('Suffix').Text
        result['bin_idx'] = win.Find('BinCombo').CurrentIndex
        result['cancelled'] = False
        disp.ExitLoop()

    def on_cancel(ev):
        disp.ExitLoop()

    def on_close(ev):
        disp.ExitLoop()

    win.On.StartBtn.Clicked = on_start
    win.On.CancelBtn.Clicked = on_cancel
    win.On.AlignWin.Close = on_close

    win.Show()
    disp.RunLoop()
    win.Hide()

    return -1 if result['cancelled'] else result['bin_idx']


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

    media_pool = project.GetMediaPool()
    aaf_timeline = project.GetCurrentTimeline()
    if not aaf_timeline:
        print("ERROR: No timeline open. Open the AAF timeline first.")
        return

    fps = float(aaf_timeline.GetSetting("timelineFrameRate"))
    print("")
    print("AAF timeline : {}".format(aaf_timeline.GetName()))
    print("Frame rate   : {} fps".format(fps))

    # Collect all bins with clips
    all_bins = []
    get_all_bins(media_pool.GetRootFolder(), all_bins)
    if not all_bins:
        print("ERROR: No bins with clips found in media pool")
        return

    # Show settings dialog or fall back to auto-detect
    fusion = get_fusion()
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
    suffix = _config['new_timeline_suffix']

    print("")
    print("Source bin   : {}".format(source_bin.GetName()))
    print("Video tracks : V1 - V{}".format(num_tracks))
    print("")
    print("Loading source timelines...")

    source_timelines = load_source_timelines(source_bin, project, fps)
    if not source_timelines:
        print("ERROR: No valid timelines found in bin.")
        print("       Make sure timelines have a valid Start Timecode set.")
        return

    # Read AAF audio track 1 to drive the alignment
    print("")
    print("Reading AAF audio track 1...")
    audio_items = aaf_timeline.GetItemListInTrack("audio", 1)
    if not audio_items:
        print("ERROR: No clips on audio track 1 of the AAF timeline")
        return
    print("Found {} audio clips".format(len(audio_items)))

    # Prepare destination timeline
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

    # Main alignment loop
    print("")
    print("Aligning...")
    print("-" * 60)

    total_placed = 0
    total_failed = 0
    total_markers = 0
    no_match_count = 0

    for audio_item in audio_items:
        clip_name = audio_item.GetName()

        if should_skip_clip(clip_name):
            print("SKIP  : {}".format(clip_name))
            continue

        mpi = audio_item.GetMediaPoolItem()
        if not mpi:
            continue

        clip_start_tc_str = mpi.GetClipProperty("Start TC")
        if not clip_start_tc_str:
            print("NO TC : {}".format(clip_name))
            continue

        clip_start_frames = tc_to_frames(clip_start_tc_str, fps)
        if clip_start_frames is None:
            continue

        left_offset = audio_item.GetLeftOffset()
        duration = audio_item.GetDuration()
        record_start = audio_item.GetStart()

        # TC range covered by this audio clip in the source material
        tc_in = clip_start_frames + left_offset
        tc_out = tc_in + duration

        src = find_source_timeline(tc_in, source_timelines)
        if not src:
            print("NO MATCH: {} (TC {})".format(clip_name, frames_to_tc(tc_in, fps)))
            no_match_count += 1
            continue

        print("MATCH : {} -> '{}' [{}]".format(
            clip_name,
            src['name'],
            frames_to_tc(tc_in, fps),
        ))

        placed, failed = copy_clips_from_source(
            src, tc_in, tc_out, record_start,
            media_pool, fps, num_tracks,
        )
        total_placed += placed
        total_failed += failed

        markers_copied = copy_markers_from_source(
            src, tc_in, tc_out, record_start, dest_timeline, fps,
        )
        total_markers += markers_copied
        if markers_copied:
            print("  + {} marker(s)".format(markers_copied))

    print("-" * 60)
    print("")
    print("=" * 60)
    print("  DONE")
    print("  Video clips placed : {}".format(total_placed))
    if total_failed:
        print("  Placement failures : {}".format(total_failed))
    print("  Markers copied     : {}".format(total_markers))
    if no_match_count:
        print("  Audio clips with no TC match : {}".format(no_match_count))
    if create_new:
        print("  Destination timeline : '{}'".format(dest_timeline.GetName()))
    print("=" * 60)
    print("")


main()
