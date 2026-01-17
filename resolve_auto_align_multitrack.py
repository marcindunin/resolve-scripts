#!/usr/bin/env python
# DaVinci Resolve - Auto Align Multitrack Clips to AAF Timeline
# ==============================================================
# Places multitrack clips on the timeline based on matching timecode
# with audio clips from an imported AAF.
#
# Author: Claude

# ============== CONFIG ==============
VIDEO_TRACK_INDEX = 1
IGNORE_PREFIXES = ["Sample", "Fade"]
MULTITRACK_BIN_NAME = "TRACKS"  # Default bin name to look for
# ====================================


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
    for prefix in IGNORE_PREFIXES:
        if clip_name.startswith(prefix):
            return True
    return False


def show_bin_selection_dialog(bins, fusion):
    """Show a dialog to select a bin using Fusion UI"""
    ui = fusion.UIManager
    disp = bmd.UIDispatcher(ui)

    # Create dropdown options
    bin_options = {}
    for i, b in enumerate(bins):
        label = "{} ({} clips)".format(b['name'], b['clip_count'])
        bin_options[i] = label

    selected_idx = [0]  # Use list to allow modification in nested function
    dialog_closed = [False]

    # Create window
    win = disp.AddWindow({
        'ID': 'BinSelector',
        'WindowTitle': 'Select Multitrack Bin',
        'Geometry': [300, 300, 400, 120],
    }, [
        ui.VGroup([
            ui.Label({'Text': 'Select bin with multitrack clips:', 'Weight': 0}),
            ui.ComboBox({'ID': 'BinCombo', 'Weight': 0}),
            ui.HGroup({'Weight': 0}, [
                ui.Button({'ID': 'OKButton', 'Text': 'OK'}),
                ui.Button({'ID': 'CancelButton', 'Text': 'Cancel'}),
            ]),
        ]),
    ])

    # Populate combo box
    combo = win.Find('BinCombo')
    for i, b in enumerate(bins):
        combo.AddItem("{} ({} clips)".format(b['name'], b['clip_count']))
    combo.CurrentIndex = 0

    def on_ok(ev):
        selected_idx[0] = win.Find('BinCombo').CurrentIndex
        dialog_closed[0] = True
        disp.ExitLoop()

    def on_cancel(ev):
        selected_idx[0] = -1
        dialog_closed[0] = True
        disp.ExitLoop()

    def on_close(ev):
        selected_idx[0] = -1
        dialog_closed[0] = True
        disp.ExitLoop()

    win.On.OKButton.Clicked = on_ok
    win.On.CancelButton.Clicked = on_cancel
    win.On.BinSelector.Close = on_close

    win.Show()
    disp.RunLoop()
    win.Hide()

    return selected_idx[0]


def show_simple_bin_list(bins):
    """Fallback: just print bins and ask user to edit CONFIG"""
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

    # Use current timeline
    timeline = project.GetCurrentTimeline()
    if not timeline:
        print("ERROR: No timeline open. Please open your AAF timeline first.")
        return

    fps = float(timeline.GetSetting("timelineFrameRate"))
    print("")
    print("Current timeline: {}".format(timeline.GetName()))
    print("Frame rate: {} fps".format(fps))

    # Get all bins with clips
    all_bins = []
    get_all_bins(root_folder, all_bins)

    if not all_bins:
        print("ERROR: No bins with clips found")
        return

    # Try to show dialog, fallback to list
    fusion = get_fusion()
    selected_idx = -1

    if fusion:
        try:
            print("")
            print("Opening bin selection dialog...")
            selected_idx = show_bin_selection_dialog(all_bins, fusion)
        except (RuntimeError, AttributeError) as e:
            print("Dialog failed: {}".format(e))
            selected_idx = -1

    if selected_idx < 0:
        # Fallback - show list and look for common names
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
            print("Please rename your multitrack bin to 'TRACKS' and run again,")
            print("or edit MULTITRACK_BIN_NAME in the script.")
            return

    multitrack_bin = all_bins[selected_idx]['folder']
    print("Selected bin: {}".format(multitrack_bin.GetName()))

    # Get multitrack clips
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

    # Get audio clips from timeline track 1
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

    # Process clips
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

    # Ensure video track exists
    while timeline.GetTrackCount("video") < VIDEO_TRACK_INDEX:
        timeline.AddTrack("video")

    # Sort and place clips
    clips_to_add.sort(key=lambda x: x['timeline_start'])

    print("")
    print("Placing {} clips on V{}...".format(len(clips_to_add), VIDEO_TRACK_INDEX))
    placed_count = 0

    for clip_info in clips_to_add:
        mc = clip_info['multitrack_clip']
        target_tc = frames_to_tc(clip_info['timeline_start'], fps)

        video_clip_info = {
            "mediaPoolItem": mc,
            "startFrame": clip_info['offset'],
            "endFrame": clip_info['offset'] + clip_info['duration'],
            "mediaType": 1,
            "trackIndex": VIDEO_TRACK_INDEX,
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
