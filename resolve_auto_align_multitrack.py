#!/usr/bin/env python
# DaVinci Resolve - Auto Align Multitrack Clips to AAF Timeline
# ==============================================================
# Places multitrack clips on the timeline based on matching timecode
# with audio clips from an imported AAF.
#
# Author: Claude

import copy

# ============== DEFAULT CONFIG ==============
DEFAULT_CONFIG = {
    'video_track_index': 1,
    'ignore_prefixes': ["Sample", "Fade"],
    'create_new_timeline': True,
    'new_timeline_suffix': '_montaz',
}
# NOTE: Audio from Multicam clips cannot be placed via the Resolve scripting API.
# The API returns success but nothing gets placed. This applies regardless of
# timeline type (AAF or normal) and whether mediaType is specified or omitted.
# ============================================

# Global config (modified by settings dialog)
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
    """Convert timecode string to frame number."""
    if not fps or fps <= 0:
        return None
    parts = timecode.replace(';', ':').split(':')
    if len(parts) != 4:
        return None
    try:
        h, m, s, f = map(int, parts)
    except ValueError:
        return None
    # Use rounded fps for frame calculation to match NLE behavior
    fps_int = int(round(fps))
    return int(((h * 3600 + m * 60 + s) * fps_int) + f)


def frames_to_tc(frames, fps):
    """Convert frame number to timecode string."""
    if not fps or fps <= 0:
        return "00:00:00:00"
    # Use rounded fps for display to match NLE behavior
    fps_int = int(round(fps))
    if fps_int <= 0:
        return "00:00:00:00"
    f = frames % fps_int
    s = (frames // fps_int) % 60
    m = (frames // (fps_int * 60)) % 60
    h = frames // (fps_int * 3600)
    return "{:02d}:{:02d}:{:02d}:{:02d}".format(h, m, s, f)


def get_clip_tc_range(clip, fps):
    start_tc = clip.GetClipProperty("Start TC")
    end_tc = clip.GetClipProperty("End TC")
    if not start_tc or not end_tc:
        return None, None
    return tc_to_frames(start_tc, fps), tc_to_frames(end_tc, fps)


def find_matching_multitrack(source_tc_frames, multitrack_clips):
    for clip_data in multitrack_clips:
        tc_start = clip_data['tc_start']
        tc_end = clip_data['tc_end']
        if tc_start is None or tc_end is None:
            continue
        if tc_start <= source_tc_frames <= tc_end:
            return clip_data
    return None


def get_all_bins(folder, bins_list, path=""):
    current_path = path + "/" + folder.GetName() if path else folder.GetName()
    clips = folder.GetClipList()
    clip_count = len(clips) if clips else 0
    if clip_count > 0:
        bins_list.append({
            'folder': folder,
            'name': folder.GetName(),
            'path': current_path,
            'clip_count': clip_count
        })
    for subfolder in folder.GetSubFolderList():
        get_all_bins(subfolder, bins_list, current_path)


def should_skip_clip(clip_name):
    for prefix in _config.get('ignore_prefixes', []):
        if clip_name.startswith(prefix):
            return True
    return False


def create_new_timeline_from_aaf(project, media_pool, aaf_timeline, suffix="_montaz"):
    """Create a new timeline with same settings as AAF timeline."""
    aaf_name = aaf_timeline.GetName()
    new_name = aaf_name + suffix

    # Get AAF timeline settings
    fps = aaf_timeline.GetSetting("timelineFrameRate")
    start_tc = aaf_timeline.GetStartTimecode()

    # Create new timeline
    new_timeline = media_pool.CreateEmptyTimeline(new_name)
    if not new_timeline:
        print("ERROR: Could not create new timeline")
        return None

    # Apply settings
    new_timeline.SetSetting("timelineFrameRate", fps)
    if start_tc:
        new_timeline.SetStartTimecode(start_tc)

    print("Created new timeline: {}".format(new_name))
    print("  Frame rate: {} fps".format(fps))
    print("  Start TC: {}".format(start_tc))

    return new_timeline


def copy_audio_from_aaf(aaf_timeline, new_timeline, media_pool, fps):
    """Copy audio clips from AAF timeline to new timeline."""
    print("")
    print("Copying audio from AAF timeline...")

    audio_track_count = aaf_timeline.GetTrackCount("audio")
    print("AAF has {} audio tracks".format(audio_track_count))

    # Ensure new timeline has enough audio tracks
    new_audio_count = new_timeline.GetTrackCount("audio")
    while new_audio_count < audio_track_count:
        new_timeline.AddTrack("audio")
        new_count = new_timeline.GetTrackCount("audio")
        if new_count == new_audio_count:
            break
        new_audio_count = new_count

    copied_count = 0

    for track_idx in range(1, audio_track_count + 1):
        items = aaf_timeline.GetItemListInTrack("audio", track_idx)
        if not items:
            continue

        for item in items:
            clip_name = item.GetName()
            if should_skip_clip(clip_name):
                continue

            media_pool_item = item.GetMediaPoolItem()
            if not media_pool_item:
                continue

            # Get clip placement info
            start_frame = item.GetStart()
            duration = item.GetDuration()
            left_offset = item.GetLeftOffset()

            clip_info = {
                "mediaPoolItem": media_pool_item,
                "startFrame": left_offset,
                "endFrame": left_offset + duration,
                "trackIndex": track_idx,
                "recordFrame": start_frame,
            }

            result = media_pool.AppendToTimeline([clip_info])
            if result:
                copied_count += 1

    print("Copied {} audio clips".format(copied_count))
    return copied_count


def show_settings_dialog(bins, fusion):
    """Show settings dialog with bin selection using Fusion UI"""
    global _config

    ui = fusion.UIManager
    disp = bmd.UIDispatcher(ui)

    result = {'bin_idx': -1, 'cancelled': True}

    win = disp.AddWindow({
        'ID': 'SettingsWin',
        'WindowTitle': 'Auto Align Multitrack - Settings',
        'Geometry': [300, 200, 450, 380],
        'Spacing': 10,
    }, [
        ui.VGroup({'Spacing': 5}, [
            ui.Label({'Text': 'Auto Align Settings', 'Font': ui.Font({'PixelSize': 16, 'Bold': True}), 'Weight': 0}),
            ui.Label({'Text': '-' * 50, 'Weight': 0}),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Target video track:', 'Weight': 2}),
                ui.SpinBox({'ID': 'VideoTrack', 'Value': _config.get('video_track_index', 1), 'Minimum': 1, 'Maximum': 10, 'Weight': 1}),
            ]),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Ignore prefixes (comma-sep):', 'Weight': 2}),
                ui.LineEdit({'ID': 'IgnorePrefixes', 'Text': ', '.join(_config.get('ignore_prefixes', [])), 'Weight': 2}),
            ]),
            ui.Label({'Text': '-' * 50, 'Weight': 0}),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Multitrack bin:', 'Weight': 1}),
                ui.ComboBox({'ID': 'BinCombo', 'Weight': 3}),
            ]),
            ui.Label({'Text': '-' * 50, 'Weight': 0}),
            ui.Label({'Text': 'New Timeline Options:', 'Font': ui.Font({'Bold': True}), 'Weight': 0}),
            ui.HGroup({'Weight': 0}, [
                ui.CheckBox({'ID': 'CreateNewTimeline', 'Text': 'Create new timeline (required for audio)', 'Checked': _config.get('create_new_timeline', True), 'Weight': 1}),
            ]),
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Timeline name suffix:', 'Weight': 2}),
                ui.LineEdit({'ID': 'TimelineSuffix', 'Text': _config.get('new_timeline_suffix', '_montaz'), 'Weight': 2}),
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
    if bins:  # Only set index if there are items
        combo.CurrentIndex = default_idx

    def on_start(ev):
        _config['video_track_index'] = win.Find('VideoTrack').Value
        prefixes_text = win.Find('IgnorePrefixes').Text
        _config['ignore_prefixes'] = [p.strip() for p in prefixes_text.split(',') if p.strip()]
        _config['create_new_timeline'] = win.Find('CreateNewTimeline').Checked
        _config['new_timeline_suffix'] = win.Find('TimelineSuffix').Text
        result['bin_idx'] = win.Find('BinCombo').CurrentIndex
        result['cancelled'] = False
        disp.ExitLoop()

    def on_cancel(ev):
        result['cancelled'] = True
        disp.ExitLoop()

    def on_close(ev):
        result['cancelled'] = True
        disp.ExitLoop()

    win.On.StartBtn.Clicked = on_start
    win.On.CancelBtn.Clicked = on_cancel
    win.On.SettingsWin.Close = on_close

    win.Show()
    disp.RunLoop()
    win.Hide()

    if result['cancelled']:
        return -1
    return result['bin_idx']


def show_simple_bin_list(bins):
    print("")
    print("Available bins with clips:")
    print("-" * 50)
    for i, b in enumerate(bins):
        print("  {}. {} ({} clips)".format(i+1, b['name'], b['clip_count']))
    print("-" * 50)
    print("")
    return None


def main():
    print("")
    print("=" * 60)
    print("  Auto Align Multitrack Clips to AAF Timeline")
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
    root_folder = media_pool.GetRootFolder()

    timeline = project.GetCurrentTimeline()
    if not timeline:
        print("ERROR: No timeline open. Please open your AAF timeline first.")
        return

    fps = float(timeline.GetSetting("timelineFrameRate"))
    print("")
    print("Current timeline: {}".format(timeline.GetName()))
    print("Frame rate: {} fps".format(fps))

    all_bins = []
    get_all_bins(root_folder, all_bins)

    if not all_bins:
        print("ERROR: No bins with clips found")
        return

    fusion = get_fusion()
    selected_idx = -1

    if fusion:
        try:
            print("")
            print("Opening settings dialog...")
            selected_idx = show_settings_dialog(all_bins, fusion)
        except (RuntimeError, AttributeError, NameError) as e:
            print("Dialog failed: {}".format(e))
            selected_idx = -1

    if selected_idx < 0:
        print("")
        print("Looking for bin named 'TRACKS'...")
        for i, b in enumerate(all_bins):
            if b['name'].upper() == 'TRACKS':
                selected_idx = i
                print("Found 'TRACKS' bin automatically!")
                break

        if selected_idx < 0:
            show_simple_bin_list(all_bins)
            print("Could not auto-detect multitrack bin.")
            print("Please rename your multitrack bin to 'TRACKS' and run again.")
            return

    # Validate selected index is in range
    if selected_idx >= len(all_bins):
        print("ERROR: Invalid bin selection")
        return

    multitrack_bin = all_bins[selected_idx]['folder']
    print("Selected bin: {}".format(multitrack_bin.GetName()))
    print("Target video track: V{}".format(_config['video_track_index']))
    print("Ignore prefixes: {}".format(_config['ignore_prefixes']))

    multitrack_clips = []
    print("")
    print("Multitrack clips found:")

    for clip in multitrack_bin.GetClipList():
        clip_name = clip.GetName()
        tc_start, tc_end = get_clip_tc_range(clip, fps)
        if tc_start is not None:
            multitrack_clips.append({
                'clip': clip,
                'name': clip_name,
                'tc_start': tc_start,
                'tc_end': tc_end
            })
            print("  - {}: {} - {}".format(clip_name, frames_to_tc(tc_start, fps), frames_to_tc(tc_end, fps)))
        else:
            print("  - {}: (no timecode - skipping)".format(clip_name))

    if not multitrack_clips:
        print("ERROR: No clips with valid timecode found in bin")
        return

    print("")
    print("Analyzing audio track 1...")

    audio_items = []
    items = timeline.GetItemListInTrack("audio", 1)
    if items:
        for item in items:
            audio_items.append({
                'item': item,
                'name': item.GetName(),
                'start': item.GetStart(),
                'end': item.GetEnd(),
                'duration': item.GetDuration(),
            })

    print("Found {} audio clips".format(len(audio_items)))

    if not audio_items:
        print("ERROR: No audio clips on track 1. Is this the AAF timeline?")
        return

    print("")
    print("Matching clips...")
    print("-" * 60)

    clips_to_add = []
    clips_skipped = 0
    clips_no_match = 0

    for audio in audio_items:
        clip_name = audio['name']

        if should_skip_clip(clip_name):
            print("SKIP: {}".format(clip_name))
            clips_skipped += 1
            continue

        audio_item = audio['item']
        media_pool_item = audio_item.GetMediaPoolItem()
        if not media_pool_item:
            clips_skipped += 1
            continue

        clip_start_tc = media_pool_item.GetClipProperty("Start TC")
        if not clip_start_tc:
            clips_skipped += 1
            continue

        clip_start_frames = tc_to_frames(clip_start_tc, fps)
        left_offset = audio_item.GetLeftOffset()
        source_in_frames = clip_start_frames + left_offset

        matching = find_matching_multitrack(source_in_frames, multitrack_clips)

        if not matching:
            print("NO MATCH: {} (TC: {})".format(clip_name, frames_to_tc(source_in_frames, fps)))
            clips_no_match += 1
            continue

        mc_start = matching['tc_start']
        offset_in_multitrack = source_in_frames - mc_start

        clips_to_add.append({
            'multitrack_clip': matching['clip'],
            'multitrack_name': matching['name'],
            'timeline_start': audio['start'],
            'duration': audio['duration'],
            'offset': offset_in_multitrack,
            'name': clip_name,
            'source_tc': frames_to_tc(source_in_frames, fps)
        })

        print("MATCH: {} -> {}".format(clip_name, matching['name']))

    print("-" * 60)
    print("")
    print("Summary:")
    print("  Clips to place: {}".format(len(clips_to_add)))
    print("  Clips skipped: {}".format(clips_skipped))
    print("  No TC match: {}".format(clips_no_match))

    if not clips_to_add:
        print("")
        print("No clips to place.")
        return

    # Get config values
    video_track_index = _config['video_track_index']
    create_new_timeline = _config.get('create_new_timeline', False)
    timeline_suffix = _config.get('new_timeline_suffix', '_montaz')

    # Store reference to AAF timeline
    aaf_timeline = timeline
    target_timeline = timeline

    # Create new timeline if requested
    if create_new_timeline:
        print("")
        print("-" * 60)
        new_timeline = create_new_timeline_from_aaf(project, media_pool, aaf_timeline, timeline_suffix)
        if not new_timeline:
            print("ERROR: Failed to create new timeline, aborting")
            return

        # Switch to new timeline
        project.SetCurrentTimeline(new_timeline)
        target_timeline = new_timeline

        # Copy audio from AAF
        copy_audio_from_aaf(aaf_timeline, new_timeline, media_pool, fps)
        print("-" * 60)

    # Ensure we have enough video tracks
    current_track_count = target_timeline.GetTrackCount("video")
    while current_track_count < video_track_index:
        if not target_timeline.AddTrack("video"):
            print("WARNING: Could not add video track")
            break
        new_count = target_timeline.GetTrackCount("video")
        if new_count == current_track_count:
            print("WARNING: Track count did not increase after AddTrack")
            break
        current_track_count = new_count

    clips_to_add.sort(key=lambda x: x['timeline_start'])

    print("")
    print("Placing {} clips on V{} (video only)...".format(len(clips_to_add), video_track_index))

    placed_count = 0

    for clip_info in clips_to_add:
        mc = clip_info['multitrack_clip']
        target_tc = frames_to_tc(clip_info['timeline_start'], fps)

        # Place video only (mediaType=1)
        # NOTE: Audio from Multicam clips cannot be placed via Resolve API
        clip_placement = {
            "mediaPoolItem": mc,
            "startFrame": clip_info['offset'],
            "endFrame": clip_info['offset'] + clip_info['duration'],
            "mediaType": 1,
            "trackIndex": video_track_index,
            "recordFrame": clip_info['timeline_start'],
        }

        result = media_pool.AppendToTimeline([clip_placement])

        if result:
            placed_count += 1
            print("  Placed: {} @ {}".format(clip_info['multitrack_name'], target_tc))
        else:
            print("  FAILED: {}".format(clip_info['name']))

    print("")
    print("=" * 60)
    print("  DONE! Placed {}/{} video clips".format(placed_count, len(clips_to_add)))
    if create_new_timeline:
        print("  New timeline: {}".format(target_timeline.GetName()))
    print("  NOTE: Audio from multitrack must be added manually")
    print("=" * 60)
    print("")


if __name__ == "__main__":
    main()
