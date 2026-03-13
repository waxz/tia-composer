"""
TIA Portal Openness — HMI Screen Manager  (corrected for V19–V21)
=================================================================
Companion module to tia_composer.py.

KEY FINDING FROM OFFICIAL DOCS (section 7.17 of the Openness System Manual)
-----------------------------------------------------------------------------
For WinCC Comfort / Advanced (HmiTarget):
  - The API does NOT have a .Create() method on ScreenComposition.
  - Screens are CREATED exclusively via XML Import (SimaticML format).
  - The documented operations on Screens are:
      • Create user-defined screen FOLDERS  → hmiTarget.ScreenFolder.Folders.Create(name)
      • Delete a screen from a folder       → screen.Delete()
      • Delete all screens from a folder    → folder.Screens.DeleteAll()  (or loop + Delete)
      • Find a screen                        → folder.Screens.Find(name)
      • Export a screen                      → screen.Export(FileInfo, ExportOptions)
      • Import a screen                      → folder.Screens.Import(FileInfo, ImportOptions)
  - There is NO programmatic way to set screen Width/Height on a fresh screen from code alone;
    those attributes live in the XML and are set during import.

For WinCC Unified (HmiUnifiedSoftware):
  - Screens are created via .Screens.Create(name)  (section 7.18.9).
  - Screen items are created via screen.ScreenItems.Create(type, name).
  - Full attribute read/write is available.

Object-model path — Comfort / Advanced
---------------------------------------
  Device
    └─ DeviceItems[3]  (varies by panel model — scan all DeviceItems)
         └─ SoftwareContainer.Software  →  HmiTarget
              └─ ScreenFolder                      (root folder)
                   ├─ Screens                      (ScreenComposition — no Create!)
                   │    └─ Screen  (find/delete/export/import)
                   └─ Folders                      (sub-folder composition)
                        └─ ScreenFolder.Create(name)

Object-model path — Unified
-----------------------------
  Device
    └─ DeviceItems[n]
         └─ SoftwareContainer.Software  →  HmiUnifiedSoftware
              └─ Screens                           (has .Create(name))
                   └─ Screen
                        └─ ScreenItems
                             └─ ScreenItem (has .Create(type, name))

Workflow for creating a Comfort screen
---------------------------------------
  1. Build a minimal SimaticML XML string (or load a template XML).
  2. Write it to a temp file.
  3. Call folder.Screens.Import(FileInfo, ImportOptions.Override).
  4. Find the imported screen via folder.Screens.Find(name).
  5. (Optionally) set additional attributes via SetAttribute.

The HmiScreenManager.create() method now follows this workflow automatically.
A built-in MINIMAL_SCREEN_XML template is used when no external XML is provided.
"""

from __future__ import annotations

import logging
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

log = logging.getLogger("tia_hmi_manager")

# ---------------------------------------------------------------------------
# Lazy .NET imports
# ---------------------------------------------------------------------------

def _tia():
    import Siemens.Engineering as m; return m          # type: ignore

def _hwf():
    import Siemens.Engineering.HW.Features as m; return m  # type: ignore

def _fi(path: str):
    from System.IO import FileInfo; return FileInfo(path)  # type: ignore

# ---------------------------------------------------------------------------
# Minimal SimaticML template for a new Comfort screen
# Placeholders: {name}, {width}, {height}, {back_color}
# ---------------------------------------------------------------------------

MINIMAL_SCREEN_XML = """\
<?xml version="1.0" encoding="utf-8"?>
<Document>
  <SW.Screens.Screen ID="0">
    <AttributeList>
      <Name>{name}</Name>
      <Width>{width}</Width>
      <Height>{height}</Height>
      <BackColor>{back_color}</BackColor>
    </AttributeList>
    <ObjectList/>
  </SW.Screens.Screen>
</Document>
"""

# ---------------------------------------------------------------------------
# Pure-data descriptors
# ---------------------------------------------------------------------------

@dataclass
class ScreenItemSpec:
    """Describes one screen object to create or update (Unified only for programmatic create)."""
    item_type: str           # e.g. "HmiButton", "HmiIOField", "HmiTextField"
    name: str
    left: int = 0
    top: int = 0
    width: int = 100
    height: int = 40
    attributes: dict[str, Any] = field(default_factory=dict)


@dataclass
class HmiScreenSpec:
    """Describes one HMI screen to create or configure."""
    name: str
    width: int = 1280
    height: int = 800
    back_color: str = "16777215"   # decimal RGB — 16777215 = white; or CSS hex "#FFFFFF"
    is_start_screen: bool = False
    comment: str = ""
    template: Optional[str] = None      # Unified only
    xml_path: Optional[str] = None      # use this XML file instead of built-in template
    items: list[ScreenItemSpec] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

def diagnose_device(device) -> None:
    """
    Print the full member list of every SoftwareContainer.Software found on a
    device, plus the member list of every Screens-like composition found on it.

    Usage:
        from tia_hmi_manager import diagnose_device
        diagnose_device(hmi_device_object)
    """
    tia = _tia(); hwf = _hwf()
    print(f"\n=== diagnose_device: {device.Name} ===")
    for idx, di in enumerate(device.DeviceItems):
        print(f"  DeviceItems[{idx}] name={di.Name!r}")
        try:
            sc = tia.IEngineeringServiceProvider(di).GetService[hwf.SoftwareContainer]()
        except Exception as e:
            print(f"    GetService failed: {e}"); continue
        if sc is None:
            print("    → None"); continue
        sw = sc.Software
        tname = type(sw).__name__
        members = sorted(m for m in dir(sw) if not m.startswith("_"))
        print(f"    Software: {tname}")
        print(f"    Members : {members}")
        for attr in ("ScreenFolder", "Screens", "ScreenGroup", "Items"):
            val = getattr(sw, attr, None)
            if val is None: continue
            sub = sorted(m for m in dir(val) if not m.startswith("_"))
            try: cnt = val.Count; print(f"    .{attr}.Count={cnt}  members={sub}")
            except: print(f"    .{attr} members={sub}")
            inner = getattr(val, "Screens", None)
            if inner is not None:
                sub2 = sorted(m for m in dir(inner) if not m.startswith("_"))
                try: cnt2 = inner.Count; print(f"      .Screens.Count={cnt2}  members={sub2}")
                except: print(f"      .Screens members={sub2}")
    print("=== end diagnose_device ===\n")


def diagnose_folder(folder) -> None:
    """Print members of an HmiTarget ScreenFolder object."""
    tname = type(folder).__name__
    members = sorted(m for m in dir(folder) if not m.startswith("_"))
    print(f"\n=== diagnose_folder ({tname}) ===")
    print(f"  Members: {members}")
    sc = getattr(folder, "Screens", None)
    if sc is not None:
        sub = sorted(m for m in dir(sc) if not m.startswith("_"))
        print(f"  .Screens members: {sub}")
    fo = getattr(folder, "Folders", None)
    if fo is not None:
        sub2 = sorted(m for m in dir(fo) if not m.startswith("_"))
        print(f"  .Folders members: {sub2}")
    print("=== end diagnose_folder ===\n")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_hmi_software(device):
    """
    Return (hmi_software, flavour_str) scanning all DeviceItems.
    flavour is 'Comfort' or 'Unified'.
    Raises RuntimeError with diagnostic hint if not found.
    """
    tia = _tia(); hwf = _hwf()
    for idx, di in enumerate(device.DeviceItems):
        try:
            sc = tia.IEngineeringServiceProvider(di).GetService[hwf.SoftwareContainer]()
        except Exception:
            continue
        if sc is None: continue
        sw = sc.Software
        tname = type(sw).__name__
        if "Hmi" not in tname: continue
        flavour = "Unified" if "Unified" in tname else "Comfort"
        log.info("Found %s (%s) on DeviceItems[%d]", tname, flavour, idx)
        return sw, flavour
    raise RuntimeError(
        f"No HMI software found on '{device.Name}'. "
        "Call diagnose_device(device) to inspect the object tree."
    )


def _get_root_folder(sw, flavour: str):
    """
    Return the root ScreenFolder (Comfort) or Screens composition (Unified).
    For Comfort, the root is sw.ScreenFolder.
    For Unified,  the root is sw.Screens  (which has .Create()).
    """
    if flavour == "Unified":
        screens = getattr(sw, "Screens", None)
        if screens is None:
            raise RuntimeError("HmiUnifiedSoftware has no .Screens attribute.")
        return screens

    # Comfort / Advanced — root is sw.ScreenFolder
    folder = getattr(sw, "ScreenFolder", None)
    if folder is not None:
        return folder

    # Some panel models name it differently
    for attr in ("ScreenGroup", "RootFolder"):
        folder = getattr(sw, attr, None)
        if folder is not None:
            return folder

    raise RuntimeError(
        "Cannot find ScreenFolder on HmiTarget. "
        "Call diagnose_device(device) to inspect the object tree."
    )


def _attr_safe_get(obj, attr: str, default=None):
    try:    return obj.GetAttribute(attr)
    except: return default


def _attr_safe_set(obj, attr: str, value) -> bool:
    try:
        obj.SetAttribute(attr, value)
        return True
    except Exception as exc:
        log.warning("SetAttribute('%s', %r) failed: %s", attr, value, exc)
        return False


def _build_screen_xml(spec: HmiScreenSpec) -> str:
    """Build a minimal SimaticML XML for one Comfort screen."""
    # Convert CSS hex to decimal if needed
    color = spec.back_color
    if color.startswith("#"):
        r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
        color = str((r << 16) | (g << 8) | b)
    return MINIMAL_SCREEN_XML.format(
        name=spec.name,
        width=spec.width,
        height=spec.height,
        back_color=color,
    )


# ---------------------------------------------------------------------------
# ScreenItemManager  (Unified only — Comfort items are edited via XML)
# ---------------------------------------------------------------------------

class ScreenItemManager:
    """
    Manage ScreenItem objects within a single Unified screen.

    For Comfort screens, items must be managed via XML export/import.
    An error is raised if you try to use add() on a Comfort screen.
    """

    def __init__(self, screen, flavour: str = "Unified"):
        self._screen = screen
        self._flavour = flavour

    def list(self) -> list[dict]:
        try:
            si = self._screen.ScreenItems
        except AttributeError:
            return []
        result = []
        for item in si:
            result.append({
                "Name":   _attr_safe_get(item, "Name"),
                "Type":   type(item).__name__,
                "Left":   _attr_safe_get(item, "Left"),
                "Top":    _attr_safe_get(item, "Top"),
                "Width":  _attr_safe_get(item, "Width"),
                "Height": _attr_safe_get(item, "Height"),
            })
        return result

    def find(self, name: str):
        try:    return self._screen.ScreenItems.Find(name)
        except: return None

    def add(self, spec: ScreenItemSpec):
        """Add a screen item. Unified only."""
        if self._flavour == "Comfort":
            raise RuntimeError(
                "Comfort HMI screen items cannot be created programmatically via Openness. "
                "Edit the screen XML and re-import it, or use the TIA Portal UI."
            )
        si = self._screen.ScreenItems
        item = si.Create(spec.item_type, spec.name)
        for attr, val in [("Left", spec.left), ("Top", spec.top),
                          ("Width", spec.width), ("Height", spec.height)]:
            _attr_safe_set(item, attr, val)
        for attr, val in spec.attributes.items():
            _attr_safe_set(item, attr, val)
        log.info("Added %s '%s'", spec.item_type, spec.name)
        return item

    def add_many(self, specs: list[ScreenItemSpec]) -> list:
        return [self.add(s) for s in specs]

    def set_attribute(self, name: str, attr: str, value) -> bool:
        item = self._require(name)
        ok = _attr_safe_set(item, attr, value)
        if ok: log.info("Set item %s.%s = %r", name, attr, value)
        return ok

    def move(self, name: str, left: int, top: int):
        item = self._require(name)
        _attr_safe_set(item, "Left", left)
        _attr_safe_set(item, "Top",  top)

    def resize(self, name: str, width: int, height: int):
        item = self._require(name)
        _attr_safe_set(item, "Width",  width)
        _attr_safe_set(item, "Height", height)

    def delete(self, name: str) -> bool:
        item = self.find(name)
        if item is None: return False
        item.Delete(); return True

    def diagnose(self) -> None:
        try:
            si = self._screen.ScreenItems
            print(sorted(m for m in dir(si) if not m.startswith("_")))
        except AttributeError:
            print("No ScreenItems attribute on this screen.")

    def _require(self, name: str):
        item = self.find(name)
        if item is None: raise KeyError(f"ScreenItem '{name}' not found.")
        return item


# ---------------------------------------------------------------------------
# HmiScreenManager
# ---------------------------------------------------------------------------

class HmiScreenManager:
    """
    Manage screens on one HMI device.

    Works for both WinCC Comfort/Advanced (HmiTarget) and WinCC Unified
    (HmiUnifiedSoftware) — but the API surface differs:

    ┌─────────────────────────┬───────────────────────┬───────────────────┐
    │ Operation               │ Comfort / Advanced    │ Unified           │
    ├─────────────────────────┼───────────────────────┼───────────────────┤
    │ create(spec)            │ XML import            │ .Screens.Create() │
    │ find(name)              │ folder.Screens.Find() │ .Screens.Find()   │
    │ delete(name)            │ screen.Delete()       │ screen.Delete()   │
    │ export(name, path)      │ screen.Export()       │ screen.Export()   │
    │ import_screen(path)     │ folder.Screens.Import │ .Screens.Import() │
    │ ScreenItem.add()        │ ✗ XML only            │ ✓                 │
    │ ScreenItem.set_attr()   │ ✓ (if supported)      │ ✓                 │
    └─────────────────────────┴───────────────────────┴───────────────────┘

    Instantiation:
        mgr = HmiScreenManager.from_device(device)

    Comfort create workflow:
        A minimal SimaticML XML is generated and imported.
        To include screen items or custom attributes, set spec.xml_path to
        a pre-built XML file (export a screen from TIA as template).
    """

    def __init__(self, sw, flavour: str, root):
        """
        sw      : HmiTarget or HmiUnifiedSoftware object
        flavour : 'Comfort' or 'Unified'
        root    : ScreenFolder (Comfort) or ScreenComposition (Unified)
        """
        self._sw = sw
        self._flavour = flavour
        self._root = root   # Comfort: ScreenFolder | Unified: Screens composition
        log.info("HmiScreenManager ready (%s)", flavour)

    @classmethod
    def from_device(cls, device) -> "HmiScreenManager":
        sw, flavour = _get_hmi_software(device)
        root = _get_root_folder(sw, flavour)
        return cls(sw, flavour, root)

    # ── internal screen composition accessor ──────────────────────────────

    @property
    def _screens(self):
        """Returns the ScreenComposition regardless of flavour."""
        if self._flavour == "Unified":
            return self._root          # IS the composition
        # Comfort: root is ScreenFolder, screens live at .Screens
        sc = getattr(self._root, "Screens", None)
        if sc is None:
            raise AttributeError(
                "ScreenFolder has no .Screens. Call diagnose_folder(mgr._root)."
            )
        return sc

    # ── query ─────────────────────────────────────────────────────────────

    def list(self) -> list[dict]:
        result = []
        for s in self._screens:
            result.append({
                "Name":   _attr_safe_get(s, "Name"),
                "Width":  _attr_safe_get(s, "Width"),
                "Height": _attr_safe_get(s, "Height"),
            })
        return result

    def find(self, name: str):
        try:    return self._screens.Find(name)
        except: return None

    def exists(self, name: str) -> bool:
        return self.find(name) is not None

    def count(self) -> int:
        try:    return self._screens.Count
        except: return sum(1 for _ in self._screens)

    def get_attribute(self, screen_name: str, attr: str):
        return _attr_safe_get(self._require(screen_name), attr)

    # ── create ────────────────────────────────────────────────────────────

    def create(self, spec: HmiScreenSpec):
        """
        Create a new screen.

        Comfort: generates a minimal SimaticML XML and imports it.
                 spec.xml_path overrides the built-in template.
        Unified: calls .Screens.Create(name) directly.
        """
        if self.exists(spec.name):
            raise ValueError(
                f"Screen '{spec.name}' already exists. "
                "Use update() or delete() it first."
            )

        if self._flavour == "Unified":
            return self._create_unified(spec)
        else:
            return self._create_comfort(spec)

    def create_or_update(self, spec: HmiScreenSpec):
        if self.exists(spec.name):
            return self.update(spec)
        return self.create(spec)

    def _create_unified(self, spec: HmiScreenSpec):
        screen = self._root.Create(spec.name)
        self._apply_spec_unified(screen, spec)
        if spec.items:
            im = ScreenItemManager(screen, "Unified")
            im.add_many(spec.items)
        log.info("Created Unified screen '%s'", spec.name)
        return screen

    def _create_comfort(self, spec: HmiScreenSpec):
        """Import a screen XML into the root ScreenFolder."""
        tia = _tia()
        if spec.xml_path:
            xml_path = spec.xml_path
            tmp = None
        else:
            xml_content = _build_screen_xml(spec)
            tmp = tempfile.NamedTemporaryFile(
                suffix=".xml", delete=False,
                mode="w", encoding="utf-8"
            )
            tmp.write(xml_content)
            tmp.close()
            xml_path = tmp.name

        try:
            self._screens.Import(
                _fi(xml_path),
                tia.ImportOptions.Override,
            )
        finally:
            if tmp:
                try: os.unlink(tmp.name)
                except: pass

        screen = self.find(spec.name)
        if screen is None:
            log.warning("Screen '%s' not found after import — XML may be malformed", spec.name)
            return None

        # Apply post-import attributes
        if spec.comment:
            _attr_safe_set(screen, "Comment", spec.comment)
        if spec.is_start_screen:
            self.set_as_start_screen(spec.name)

        log.info("Created Comfort screen '%s' via XML import", spec.name)
        return screen

    # ── update ────────────────────────────────────────────────────────────

    def update(self, spec: HmiScreenSpec):
        """
        Update an existing screen.
        Comfort: re-imports the XML (Override replaces the existing screen).
        Unified: patches attributes directly.
        """
        screen = self._require(spec.name)
        if self._flavour == "Unified":
            self._apply_spec_unified(screen, spec)
            if spec.items:
                ScreenItemManager(screen, "Unified").add_many(spec.items)
        else:
            # For Comfort, re-import is the only way to change screen properties
            self._create_comfort(spec)
        log.info("Updated screen '%s'", spec.name)
        return self.find(spec.name)

    def set_attribute(self, screen_name: str, attr: str, value) -> bool:
        screen = self._require(screen_name)
        ok = _attr_safe_set(screen, attr, value)
        if ok: log.info("Screen '%s': %s = %r", screen_name, attr, value)
        return ok

    def set_attributes(self, screen_name: str, attrs: dict) -> dict:
        screen = self._require(screen_name)
        return {a: _attr_safe_set(screen, a, v) for a, v in attrs.items()}

    def rename(self, old_name: str, new_name: str) -> None:
        screen = self._require(old_name)
        try:    screen.Rename(new_name)
        except AttributeError:
            _attr_safe_set(screen, "Name", new_name)
        log.info("Renamed '%s' → '%s'", old_name, new_name)

    def resize(self, screen_name: str, width: int, height: int) -> None:
        s = self._require(screen_name)
        _attr_safe_set(s, "Width", width)
        _attr_safe_set(s, "Height", height)
        log.info("Resized '%s' to %dx%d", screen_name, width, height)

    def set_background(self, screen_name: str, color) -> None:
        """color: CSS hex '#RRGGBB' or decimal int."""
        if isinstance(color, str) and color.startswith("#"):
            r,g,b = int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
            color = (r << 16) | (g << 8) | b
        self.set_attribute(screen_name, "BackColor", color)

    def set_as_start_screen(self, screen_name: str) -> None:
        for s in self._screens:
            _attr_safe_set(s, "IsStartScreen", False)
        _attr_safe_set(self._require(screen_name), "IsStartScreen", True)
        log.info("'%s' set as start screen", screen_name)

    # ── delete ────────────────────────────────────────────────────────────

    def delete(self, screen_name: str) -> bool:
        screen = self.find(screen_name)
        if screen is None:
            log.warning("Screen '%s' not found — skipped", screen_name)
            return False
        screen.Delete()
        log.info("Deleted screen '%s'", screen_name)
        return True

    def delete_all(self) -> int:
        names = [_attr_safe_get(s, "Name") for s in self._screens]
        for name in (n for n in names if n):
            self.delete(name)
        return len(names)

    # ── export / import ───────────────────────────────────────────────────

    def export(self, screen_name: str, out_path: str, options=None) -> Path:
        screen = self._require(screen_name)
        p = Path(out_path)
        if p.is_dir() or not p.suffix:
            p = p / f"{screen_name}.xml"
        p.parent.mkdir(parents=True, exist_ok=True)
        opts = options if options is not None else _tia().ExportOptions.WithDefaults
        screen.Export(_fi(str(p)), opts)
        log.info("Exported '%s' → %s", screen_name, p)
        return p

    def import_screen(self, xml_path: str, options=None) -> None:
        p = Path(xml_path)
        if not p.exists():
            raise FileNotFoundError(xml_path)
        opts = options if options is not None else _tia().ImportOptions.Override
        self._screens.Import(_fi(str(p)), opts)
        log.info("Imported screen from %s", p)

    def export_all(self, out_dir: str, options=None) -> list[Path]:
        paths = []
        for s in self._screens:
            name = _attr_safe_get(s, "Name")
            if name:
                paths.append(self.export(name, out_dir, options))
        return paths

    def import_all(self, xml_dir: str, options=None) -> int:
        count = 0
        for p in Path(xml_dir).glob("*.xml"):
            self.import_screen(str(p), options)
            count += 1
        return count

    def clone(self, source_name: str, new_name: str, export_dir: str = None) -> None:
        """Export source, patch the Name in the XML, re-import as new_name."""
        import re
        tmp_dir = export_dir or tempfile.mkdtemp()
        xml_path = self.export(source_name, tmp_dir)
        text = xml_path.read_text(encoding="utf-8")
        patched = re.sub(
            r'(<Name>)[^<]*(</Name>)',
            lambda m: m.group(1) + new_name + m.group(2),
            text, count=1,
        )
        out = Path(tmp_dir) / f"{new_name}.xml"
        out.write_text(patched, encoding="utf-8")
        self.import_screen(str(out))
        log.info("Cloned '%s' → '%s'", source_name, new_name)

    # ── folder management (Comfort only) ─────────────────────────────────

    def create_folder(self, folder_name: str):
        """Create a sub-folder under the root ScreenFolder (Comfort only)."""
        if self._flavour == "Unified":
            log.warning("Folder management is a Comfort-only feature.")
            return None
        folders = getattr(self._root, "Folders", None)
        if folders is None:
            raise AttributeError(
                "ScreenFolder has no .Folders composition. "
                "Call diagnose_folder(mgr._root) to inspect."
            )
        folder = folders.Create(folder_name)
        log.info("Created screen folder '%s'", folder_name)
        return folder

    def get_folder(self, folder_name: str):
        """Find a named sub-folder (Comfort only)."""
        if self._flavour == "Unified":
            return None
        folders = getattr(self._root, "Folders", None)
        if folders is None:
            return None
        return folders.Find(folder_name)

    # ── item access ───────────────────────────────────────────────────────

    def items(self, screen) -> ScreenItemManager:
        return ScreenItemManager(screen, self._flavour)

    def items_by_name(self, screen_name: str) -> ScreenItemManager:
        return ScreenItemManager(self._require(screen_name), self._flavour)

    # ── private ───────────────────────────────────────────────────────────

    def _require(self, name: str):
        s = self.find(name)
        if s is None:
            available = [_attr_safe_get(x, "Name") for x in self._screens]
            raise KeyError(f"Screen '{name}' not found. Available: {available}")
        return s

    def _apply_spec_unified(self, screen, spec: HmiScreenSpec) -> None:
        """Apply HmiScreenSpec fields to a Unified screen object."""
        _attr_safe_set(screen, "Width",  spec.width)
        _attr_safe_set(screen, "Height", spec.height)
        color = spec.back_color
        if isinstance(color, str) and color.startswith("#"):
            r,g,b = int(color[1:3],16),int(color[3:5],16),int(color[5:7],16)
            color = (r << 16) | (g << 8) | b
        _attr_safe_set(screen, "BackColor", color)
        if spec.comment:
            _attr_safe_set(screen, "Comment", spec.comment)
        if spec.template:
            _attr_safe_set(screen, "TemplateScreenName", spec.template)
        if spec.is_start_screen:
            self.set_as_start_screen(spec.name)


# ---------------------------------------------------------------------------
# ProjectComposer integration helper
# ---------------------------------------------------------------------------

def get_hmi_manager_from_composer(composer, hmi_device_name: str) -> HmiScreenManager:
    """
    Convenience: get an HmiScreenManager from a ProjectComposer instance.

    Usage:
        hmi = get_hmi_manager_from_composer(composer, "HMI1")
        hmi.create(HmiScreenSpec("MainScreen", width=1280, height=800))
    """
    device = composer._devices.get(hmi_device_name)
    if device is None:
        raise KeyError(
            f"Device '{hmi_device_name}' not in composer. "
            f"Available: {list(composer._devices)}"
        )
    return HmiScreenManager.from_device(device)


# ---------------------------------------------------------------------------
# Example usage
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    VERSION = os.environ.get("TIA_VERSION", "V21")
    DLL = (
        rf"C:\Program Files\Siemens\Automation\Portal {VERSION}"
        rf"\PublicAPI\{VERSION}\net48\Siemens.Engineering.Step7.dll"
    )
    try:
        import clr; clr.AddReference(DLL)  # type: ignore
    except Exception as e:
        print(f"Cannot load DLL: {e}"); sys.exit(1)

    import Siemens.Engineering as tia  # type: ignore
    processes = tia.TiaPortal.GetProcesses()
    if not processes:
        print("No running TIA Portal found."); sys.exit(1)

    portal  = processes[0].Attach()
    project = portal.Projects[0]

    hmi_device = next(
        (d for d in project.Devices if "HMI" in d.Name.upper()), None
    )
    if hmi_device is None:
        print("No HMI device found."); sys.exit(1)

    # -- Uncomment to inspect object tree if you hit AttributeError:
    # diagnose_device(hmi_device)

    mgr = HmiScreenManager.from_device(hmi_device)

    # -- Uncomment to inspect ScreenFolder members:
    # diagnose_folder(mgr._root)

    # Create a screen (Comfort: via XML import; Unified: via .Create())
    mgr.create(HmiScreenSpec(
        name="MainScreen",
        width=1280,
        height=800,
        back_color="#1A1A2E",
        is_start_screen=True,
        comment="Main overview",
    ))

    # Create another, then clone it
    mgr.create(HmiScreenSpec("DetailScreen", width=1280, height=800))
    mgr.clone("DetailScreen", "DetailScreen_Line2")

    # Export all screens
    exported = mgr.export_all(r"C:\TIA\exports\screens")
    print(f"Exported {len(exported)} screens")

    # List
    for info in mgr.list():
        print(info)

    # Delete one
    mgr.delete("DetailScreen_Line2")

    project.Save()
    print("Done.")