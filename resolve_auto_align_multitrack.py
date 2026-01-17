#!/usr/bin/env python
# DaVinci Resolve - Auto Align Multitrack Clips to AAF Timeline
# ==============================================================
# Places multitrack clips on the timeline based on matching timecode
# with audio clips from an imported AAF.
#
# Author: Claude

# ============== DEFAULT CONFIG ==============
DEFAULT_CONFIG = {
    'video_track_index': 1,
    'ignore_prefixes': ["Sample", "Fade"],
}
# ============================================

# Global config (modified by settings dialog)
_config = DEFAULT_CONFIG.copy()


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
    parts = timecode.replace(';', ':').split(':')
    if len(parts) != 4:
        return None
    h, m, s, f = map(int, parts)
    return int(((h * 3600 + m * 60 + s) * fps) + f)


def frames_to_tc(frames, fps):
    fps = int(round(fps))
    f = frames % fps
    s = (frames // fps) % 60
    m = (frames // (fps * 60)) % 60
    h = frames // (fps * 3600)
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


def show_settings_dialog(bins, fusion):
    """Show settings dialog with bin selection using Fusion UI"""
    global _config

    ui = fusion.UIManager
    disp = bmd.UIDispatcher(ui)

    result = {'bin_idx': -1, 'cancelled': True}

    win = disp.AddWindow({
        'ID': 'SettingsWin',
        'WindowTitle': 'Auto Align Multitrack - Settings',
        'Geometry': [300, 200, 450, 220],
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
    combo.CurrentIndex = default_idx

    def on_start(ev):
        _config['video_track_index'] = win.Find('VideoTrack').Value
        prefixes_text = win.Find('IgnorePrefixes').Text
        _config['ignore_prefixes'] = [p.strip() for p in prefixes_text.split(',') if p.strip()]
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

    video_track_index = _config['video_track_index']

    while timeline.GetTrackCount("video") < video_track_index:
        timeline.AddTrack("video")

    clips_to_add.sort(key=lambda x: x['timeline_start'])

    print("")
    print("Placing {} clips on V{}...".format(len(clips_to_add), video_track_index))
    placed_count = 0

    for clip_info in clips_to_add:
        mc = clip_info['multitrack_clip']
        target_tc = frames_to_tc(clip_info['timeline_start'], fps)

        video_clip_info = {
            "mediaPoolItem": mc,
            "startFrame": clip_info['offset'],
            "endFrame": clip_info['offset'] + clip_info['duration'],
            "mediaType": 1,
            "trackIndex": video_track_index,
            "recordFrame": clip_info['timeline_start'],
        }

        result = media_pool.AppendToTimeline([video_clip_info])

        if result:
            placed_count += 1
            print("  Placed: {} @ {}".format(clip_info['multitrack_name'], target_tc))
        else:
            print("  FAILED: {}".format(clip_info['name']))

    print("")
    print("=" * 60)
    print("  DONE! Placed {}/{} clips".format(placed_count, len(clips_to_add)))
    print("=" * 60)
    print("")


if __name__ == "__main__":
    main()
