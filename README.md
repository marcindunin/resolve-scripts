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

## License

MIT License - feel free to use and modify.
