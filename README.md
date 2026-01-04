# DaVinci Resolve Scripts

A collection of Python scripts for DaVinci Resolve automation.

## Scripts

### resolve_auto_align_multitrack.py

Automatically aligns multitrack video clips to an AAF timeline based on matching timecode.

**Use case:** When you have an AAF from audio post (Pro Tools, etc.) with audio clips, and you want to automatically place the corresponding multitrack video files at the correct timeline positions.

#### How it works

1. Reads audio clips from track 1 of the current timeline (your imported AAF)
2. Matches each audio clip's source timecode to multitrack clips in a selected bin
3. Places the matching video clips on V1 at the correct timeline positions

#### Installation

**macOS:**
```
~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/
```

**Windows:**
```
%APPDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\
```

**Linux:**
```
~/.local/share/DaVinciResolve/Fusion/Scripts/Utility/
```

#### Usage

1. Import your AAF into DaVinci Resolve
2. Import your multitrack video clips into a bin (e.g., named "TRACKS")
3. Open the AAF timeline
4. Run the script from: **Workspace > Scripts > resolve_auto_align_multitrack**
5. Select the bin containing your multitrack clips from the dropdown dialog
6. The script will match and place clips automatically

#### Configuration

Edit the top of the script to customize:

```python
VIDEO_TRACK_INDEX = 1          # Target video track for placed clips
IGNORE_PREFIXES = ["Sample", "Fade"]  # Skip clips starting with these prefixes
```

#### Requirements

- DaVinci Resolve 18+ (tested on Resolve 20)
- Multitrack clips must have valid timecode matching the AAF source clips

---

### resolve_timeline_qc.py

Quality Control script with full GUI that checks your timeline for common issues.

**Use case:** Before final export, run this script to catch blank frames, audio overlaps, flash frames, and other potential issues.

#### Checks performed

- **Video gaps** - Frames with no video on any track
- **Flash frames** - Very short clips (< 3 frames by default), filters out AAF artifacts
- **Audio overlaps** - Two clips overlapping on the same audio track
- **Audio gaps** - Missing audio between clips
- **Disabled/muted clips** - Clips or tracks that may have been accidentally disabled
- **Offline media** - Verifies source files exist on disk using file system check
- **Clips at source end** - Clips trimmed to the very end of source (disabled by default)

#### Features

- **Full GUI interface** with three windows:
  - **Settings Window** - Configure all options, save settings, start QC
  - **Progress Window** - Real-time console output during analysis
  - **Results Window** - Interactive issue list with navigation
- **Click-to-jump** - Double-click any issue to jump to that timecode
- **Navigation** - Previous/Next buttons to step through issues
- **Export Report** - Save report to Desktop as text file
- **Persistent settings** - Configuration saved to `timeline_qc_config.json`
- Automatically skips adjustment clips in all checks
- Filters out AAF artifacts ("Sample accurate edit", "Fade" clips)
- Falls back to console mode if GUI unavailable

#### Usage

1. Open the timeline you want to check
2. Run the script from: **Workspace > Scripts > resolve_timeline_qc**
3. Configure settings in the Settings window
4. Click **Start QC** to begin analysis
5. Review progress in the Progress window
6. Browse results and click issues to jump to them on the timeline

#### Settings

All settings can be configured in the GUI:

| Setting | Description | Default |
|---------|-------------|---------|
| Flash frame threshold | Clips shorter than this are flagged | 3 frames |
| Min audio gap | Ignore gaps smaller than this | 2 frames |
| Ignore prefixes | Skip clips starting with these (comma-separated) | Sample, Fade |
| Check audio gaps | Enable/disable audio gap detection | On |
| Check offline media | Enable/disable offline media detection | On |
| Check clips at source end | Enable/disable source end detection | Off |
| Ignore adjustment clips | Skip adjustment clips in all checks | On |

#### Known limitations

- **Clip fade handles** - The Resolve scripting API does not expose clip fade handles (the fades you create by dragging clip corners). These cannot be detected by any script.

---

## Installation

Copy any script to the Resolve scripts folder:

**macOS:**
```
~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/
```

**Windows:**
```
%APPDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\
```

**Linux:**
```
~/.local/share/DaVinciResolve/Fusion/Scripts/Utility/
```

Then access via **Workspace > Scripts** in DaVinci Resolve.

## License

MIT License - feel free to use and modify.
