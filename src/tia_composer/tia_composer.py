"""
TIA Portal Openness — Modular PLC Composer
=========================================
Architecture:
  - TiaSession      : lifecycle manager (connect / attach / dispose)
  - DeviceSpec      : pure-data descriptor for any hardware item (MLFB + slot + name)
  - NetworkSpec     : pure-data descriptor for IP / subnet / IO-system wiring
  - SoftwareSpec    : pure-data descriptor for blocks / tag-tables to import
  - ProjectComposer : orchestrates DeviceSpec → DeviceSpec → NetworkSpec → SoftwareSpec
  - BlockManager    : import / export / delete blocks
  - TagManager      : import / export / delete tag tables
  - CompileManager  : HW or SW compile with structured result logging

Usage
-----
  from tia_composer import TiaSession, DeviceSpec, NetworkSpec, ProjectComposer

  specs = [
      DeviceSpec("PLC1",    "OrderNumber:6ES7 513-1AL02-0AB0/V2.6",  None, ip="192.168.0.130"),
      DeviceSpec("IOnode1", "OrderNumber:6ES7 155-6AU01-0BN0/V4.1",  None, ip="192.168.0.131"),
      DeviceSpec("HMI1",    "OrderNumber:6AV2 124-0MC01-0AX0/17.0.0.0", None),
  ]
  io_cards = [
      DeviceSpec("IO1", "OrderNumber:6ES7 521-1BL00-0AB0/V2.1", slot=2, parent="PLC1"),
      DeviceSpec("IO1", "OrderNumber:6ES7 131-6BH01-0BA0/V0.0", slot=1, parent="IOnode1"),
  ]
  net = NetworkSpec(subnet_name="Profinet", io_system_name="PNIO")

  with TiaSession(ui=True) as session:
      composer = ProjectComposer(session, project_dir=r"C:\TIA", project_name="ModularDemo")
      composer.build(specs, io_cards, net)
      composer.compile_all_hw()
      composer.compile_all_sw()
      composer.save()
"""

from __future__ import annotations

import os
import sys
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

log = logging.getLogger("tia_composer")
logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(name)s  %(message)s")

# ---------------------------------------------------------------------------
# Bootstrap pythonnet + Siemens DLL
# ---------------------------------------------------------------------------

VERSION = os.environ.get("TIA_VERSION", "V21")
_DLL = (
    rf"C:\Program Files\Siemens\Automation\Portal {VERSION}"
    rf"\PublicAPI\{VERSION}\net48\Siemens.Engineering.Step7.dll"
)

def _load_clr(dll_path: str = _DLL) -> None:
    try:
        import clr  # type: ignore
        clr.AddReference(dll_path)
        log.info("Loaded Siemens DLL: %s", dll_path)
    except ImportError:
        log.error("pythonnet not installed — pip install pythonnet")
        sys.exit(1)
    except Exception as exc:
        log.error("Cannot load DLL %s: %s", dll_path, exc)
        sys.exit(1)

_load_clr()

from System.IO import DirectoryInfo, FileInfo  # type: ignore
import Siemens.Engineering as tia              # type: ignore
import Siemens.Engineering.HW.Features as hwf  # type: ignore
import Siemens.Engineering.Compiler as comp    # type: ignore


# ---------------------------------------------------------------------------
# Pure-data descriptors (no TIA objects here)
# ---------------------------------------------------------------------------

@dataclass
class DeviceSpec:
    """Describes one physical device or an IO-card to be plugged into a device."""
    name: str
    mlfb: str                          # OrderNumber string
    slot: Optional[int] = None         # None → top-level device; int → plug into parent
    parent: Optional[str] = None       # parent DeviceSpec.name when slot is set
    ip: Optional[str] = None           # IPv4 string; None → skip network assignment
    display_name: Optional[str] = None # TIA display name (defaults to name)
    # Runtime reference filled by ProjectComposer
    _device_object: object = field(default=None, repr=False, compare=False)

    def __post_init__(self):
        if self.display_name is None:
            self.display_name = self.name


@dataclass
class NetworkSpec:
    """Describes the PROFINET subnet + IO system to create."""
    subnet_name: str = "Profinet"
    io_system_name: str = "PNIO"
    # Controller is assumed to be the device whose network interface is index 0
    # after scanning. Set controller_name to force a specific device name.
    controller_name: Optional[str] = None


@dataclass
class SoftwareSpec:
    """Describes blocks / tag tables to import into a named PLC."""
    plc_name: str
    blocks: list[str] = field(default_factory=list)       # paths to XML block files
    tag_tables: list[str] = field(default_factory=list)   # paths to XML tag-table files


# ---------------------------------------------------------------------------
# TiaSession — lifecycle / context manager
# ---------------------------------------------------------------------------

class TiaSession:
    """
    Context manager that owns the TIA Portal process.

    with TiaSession(ui=True) as session:
        project = session.create_project(...)
        ...
    # session.dispose() called automatically

    Or attach to an already-running TIA:
        with TiaSession.attach() as session: ...
    """

    def __init__(self, ui: bool = True, dll_path: str = _DLL):
        self._ui = ui
        self._portal: tia.TiaPortal | None = None
        self.project = None

    # -- context manager --

    def __enter__(self) -> "TiaSession":
        mode = tia.TiaPortalMode.WithUserInterface if self._ui else tia.TiaPortalMode.WithoutUserInterface
        log.info("Starting TIA Portal (ui=%s)…", self._ui)
        self._portal = tia.TiaPortal(mode)
        return self

    def __exit__(self, *_):
        self.dispose()

    # -- alternative: attach to existing process --

    @classmethod
    def attach(cls, index: int = 0) -> "TiaSession":
        inst = cls.__new__(cls)
        processes = tia.TiaPortal.GetProcesses()
        if not processes:
            raise RuntimeError("No running TIA Portal processes found.")
        inst._portal = processes[index].Attach()
        inst.project = inst._portal.Projects[0]
        log.info("Attached to existing TIA process, project: %s", inst.project.Name)
        return inst

    def dispose(self):
        if self._portal is not None:
            log.info("Disposing TIA Portal…")
            self._portal.Dispose()
            self._portal = None

    # -- project helpers --

    def create_project(self, directory: str, name: str) -> object:
        final = os.path.join(directory, name)
        if os.path.exists(final):
            raise FileExistsError(
                f"Project directory already exists: {final}\n"
                "Remove it or choose a different name."
            )
        d = DirectoryInfo(directory)
        self.project = self._portal.Projects.Create(d, name)
        log.info("Created project: %s", final)
        return self.project

    def open_project(self, path: str) -> object:
        self.project = self._portal.Projects.OpenWithUpgrade(FileInfo(path))
        log.info("Opened project: %s", path)
        return self.project

    def save_project(self):
        if self.project:
            self.project.Save()
            log.info("Project saved.")


# ---------------------------------------------------------------------------
# Helpers for network scanning
# ---------------------------------------------------------------------------

def _collect_network_interfaces(project) -> list:
    """Return all NetworkInterface objects found across all devices."""
    interfaces = []
    for device in project.Devices:
        for di in device.DeviceItems[1].DeviceItems:
            svc = tia.IEngineeringServiceProvider(di).GetService[hwf.NetworkInterface]()
            if isinstance(svc, hwf.NetworkInterface):
                interfaces.append((device.Name, svc))
    return interfaces  # list of (device_name, NetworkInterface)


# ---------------------------------------------------------------------------
# BlockManager
# ---------------------------------------------------------------------------

class BlockManager:
    """Import / export / delete PLC software blocks."""

    def __init__(self, software_base):
        self._sw = software_base

    def export(self, block_name: str, out_path: str,
               options=tia.ExportOptions.WithDefaults) -> Path:
        blk = self._sw.BlockGroup.Blocks.Find(block_name)
        if blk is None:
            raise KeyError(f"Block not found: {block_name!r}")
        p = Path(out_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        blk.Export(FileInfo(str(p)), options)
        log.info("Exported block '%s' → %s", block_name, p)
        return p

    def import_block(self, xml_path: str,
                     options=tia.ImportOptions.Override):
        p = Path(xml_path)
        if not p.exists():
            raise FileNotFoundError(xml_path)
        self._sw.BlockGroup.Blocks.Import(FileInfo(str(p)), options)
        log.info("Imported block from %s", p)

    def delete(self, block_name: str):
        blk = self._sw.BlockGroup.Blocks.Find(block_name)
        if blk:
            blk.Delete()
            log.info("Deleted block '%s'", block_name)


# ---------------------------------------------------------------------------
# TagManager
# ---------------------------------------------------------------------------

class TagManager:
    """Import / export / delete PLC tag tables."""

    def __init__(self, software_base):
        self._grp = software_base.TagTableGroup

    def export(self, table_name: str, out_path: str,
               options=tia.ExportOptions.WithDefaults) -> Path:
        tbl = self._grp.TagTables.Find(table_name)
        if tbl is None:
            raise KeyError(f"Tag table not found: {table_name!r}")
        p = Path(out_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        tbl.Export(FileInfo(str(p)), options)
        log.info("Exported tag table '%s' → %s", table_name, p)
        return p

    def import_table(self, xml_path: str,
                     options=tia.ImportOptions.Override):
        p = Path(xml_path)
        if not p.exists():
            raise FileNotFoundError(xml_path)
        self._grp.TagTables.Import(FileInfo(str(p)), options)
        log.info("Imported tag table from %s", p)

    def create(self, name: str):
        self._grp.TagTables.Create(name)
        log.info("Created tag table '%s'", name)

    def delete(self, name: str):
        tbl = self._grp.TagTables.Find(name)
        if tbl:
            tbl.Delete()
            log.info("Deleted tag table '%s'", name)


# ---------------------------------------------------------------------------
# CompileManager
# ---------------------------------------------------------------------------

class CompileManager:
    """Compile HW or SW and return a structured result."""

    @staticmethod
    def _print_messages(messages, depth: int = 0):
        indent = "  " * depth
        for msg in messages:
            log.info("%sPath: %s", indent, msg.Path)
            log.info("%sState: %s | Warn: %s | Err: %s | %s",
                     indent, msg.State, msg.WarningCount,
                     msg.ErrorCount, msg.Description)
            CompileManager._print_messages(msg.Messages, depth + 1)

    @staticmethod
    def compile_device_hw(device) -> dict:
        """Full HW+SW compile for one device."""
        svc = device.GetService[comp.ICompilable]()
        result = svc.Compile()
        log.info("[HW] %s — state=%s warn=%s err=%s",
                 device.Name, result.State, result.WarningCount, result.ErrorCount)
        CompileManager._print_messages(result.Messages)
        return {"state": str(result.State), "warnings": result.WarningCount,
                "errors": result.ErrorCount}

    @staticmethod
    def compile_device_sw(device, project) -> dict:
        """SW-only compile for one device (iterates DeviceItems)."""
        results = []
        for di in device.DeviceItems:
            sc = tia.IEngineeringServiceProvider(di).GetService[hwf.SoftwareContainer]()
            if sc is not None:
                sw = sc.Software
                svc = sw.GetService[comp.ICompilable]()
                result = svc.Compile()
                log.info("[SW] %s — state=%s warn=%s err=%s",
                         di.Name, result.State, result.WarningCount, result.ErrorCount)
                CompileManager._print_messages(result.Messages)
                results.append({
                    "device_item": di.Name,
                    "state": str(result.State),
                    "warnings": result.WarningCount,
                    "errors": result.ErrorCount,
                })
        return results


# ---------------------------------------------------------------------------
# ProjectComposer — the high-level orchestrator
# ---------------------------------------------------------------------------

class ProjectComposer:
    """
    Orchestrates creation of a TIA project from declarative specs.

    Workflow:
        composer = ProjectComposer(session, project_dir, project_name)
        composer.build(device_specs, io_card_specs, network_spec)
        composer.apply_software(software_specs)  # optional
        composer.compile_all_hw()
        composer.compile_all_sw()
        composer.save()
    """

    def __init__(self, session: TiaSession, project_dir: str, project_name: str):
        self._session = session
        self._project = session.create_project(project_dir, project_name)
        # name → TIA device object
        self._devices: dict[str, object] = {}

    # -- public API --

    def build(self,
              device_specs: list[DeviceSpec],
              io_card_specs: list[DeviceSpec],
              network_spec: NetworkSpec) -> None:
        """Full build: devices → IO cards → network."""
        self._add_devices(device_specs)
        self._add_io_cards(io_card_specs)
        self._configure_network(device_specs, network_spec)

    def apply_software(self, specs: list[SoftwareSpec]) -> None:
        for spec in specs:
            device = self._require_device(spec.plc_name)
            sw = self._get_software(device)
            bm, tm = BlockManager(sw), TagManager(sw)
            for path in spec.tag_tables:
                tm.import_table(path)
            for path in spec.blocks:
                bm.import_block(path)

    def compile_all_hw(self) -> list[dict]:
        results = []
        for name, dev in self._devices.items():
            results.append({name: CompileManager.compile_device_hw(dev)})
        return results

    def compile_all_sw(self) -> list[dict]:
        results = []
        for name, dev in self._devices.items():
            results.append({name: CompileManager.compile_device_sw(dev, self._project)})
        return results

    def save(self):
        self._session.save_project()

    # -- convenience accessors --

    def get_block_manager(self, plc_name: str) -> BlockManager:
        return BlockManager(self._get_software(self._require_device(plc_name)))

    def get_tag_manager(self, plc_name: str) -> TagManager:
        return TagManager(self._get_software(self._require_device(plc_name)))

    # -- private helpers --

    def _add_devices(self, specs: list[DeviceSpec]) -> None:
        for spec in specs:
            # HMI devices use None as the device name (TIA convention)
            tia_name = None if "HMI" in spec.name.upper() else spec.name
            dev = self._project.Devices.CreateWithItem(spec.mlfb, spec.display_name, tia_name)
            self._devices[spec.name] = dev
            spec._device_object = dev
            log.info("Created device: %s (%s)", spec.name, spec.mlfb)

    def _add_io_cards(self, specs: list[DeviceSpec]) -> None:
        for spec in specs:
            if spec.parent is None or spec.slot is None:
                log.warning("IO-card spec '%s' missing parent/slot — skipped", spec.name)
                continue
            parent_dev = self._require_device(spec.parent)
            di = parent_dev.DeviceItems[0]
            if di.CanPlugNew(spec.mlfb, spec.name, spec.slot):
                di.PlugNew(spec.mlfb, spec.name, spec.slot)
                log.info("Plugged %s into %s slot %d", spec.name, spec.parent, spec.slot)
            else:
                log.warning("Cannot plug %s into %s slot %d", spec.name, spec.parent, spec.slot)

    def _configure_network(self, device_specs: list[DeviceSpec],
                           net: NetworkSpec) -> None:
        # Collect all interfaces in project order
        ifaces = _collect_network_interfaces(self._project)  # [(dev_name, iface)]

        # Apply IPs from DeviceSpec lookup
        ip_map = {s.name: s.ip for s in device_specs if s.ip}
        for dev_name, iface in ifaces:
            ip = ip_map.get(dev_name)
            if ip:
                iface.Nodes[0].SetAttribute("Address", ip)
                log.info("Assigned IP %s → %s", ip, dev_name)

        # Determine controller (first in list, or by name)
        controller_iface = None
        io_system = None
        subnet = None

        for idx, (dev_name, iface) in enumerate(ifaces):
            is_controller = (
                net.controller_name is None and idx == 0
            ) or (dev_name == net.controller_name)

            if is_controller:
                subnet = iface.Nodes[0].CreateAndConnectToSubnet(net.subnet_name)
                io_system = iface.IoControllers[0].CreateIoSystem(net.io_system_name)
                controller_iface = iface
                log.info("Created subnet '%s' and IO system '%s' on %s",
                         net.subnet_name, net.io_system_name, dev_name)
            else:
                iface.Nodes[0].ConnectToSubnet(subnet)
                if iface.IoConnectors.Count > 0:
                    iface.IoConnectors[0].ConnectToIoSystem(io_system)
                log.info("Connected %s to subnet/IO system", dev_name)

    def _require_device(self, name: str):
        if name not in self._devices:
            raise KeyError(f"Device '{name}' not found. Available: {list(self._devices)}")
        return self._devices[name]

    @staticmethod
    def _get_software(device) -> object:
        """Return the PlcSoftware object from a device (iterates DeviceItems)."""
        for di in device.DeviceItems:
            sc = tia.IEngineeringServiceProvider(di).GetService[hwf.SoftwareContainer]()
            if sc is not None:
                return sc.Software
        raise RuntimeError(f"No software container found on {device.Name}")


# ---------------------------------------------------------------------------
# Example usage (guarded so it only runs when executed directly)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOME = os.path.expanduser("~")

    device_specs = [
        DeviceSpec("PLC1",    "OrderNumber:6ES7 513-1AL02-0AB0/V2.6",   ip="192.168.0.130"),
        DeviceSpec("IOnode1", "OrderNumber:6ES7 155-6AU01-0BN0/V4.1",   ip="192.168.0.131"),
        DeviceSpec("HMI1",    "OrderNumber:6AV2 124-0MC01-0AX0/17.0.0.0"),
    ]

    io_card_specs = [
        DeviceSpec("IO1", "OrderNumber:6ES7 521-1BL00-0AB0/V2.1", slot=2, parent="PLC1"),
        DeviceSpec("IO1", "OrderNumber:6ES7 131-6BH01-0BA0/V0.0", slot=1, parent="IOnode1"),
    ]

    net = NetworkSpec(subnet_name="Profinet", io_system_name="PNIO", controller_name="PLC1")

    # Optional: import existing XML exports
    sw_specs = [
        SoftwareSpec(
            plc_name="PLC1",
            blocks=[r"C:\TIA\exports\Main.xml"],
            tag_tables=[r"C:\TIA\exports\dummy.xml"],
        )
    ]

    with TiaSession(ui=True) as session:
        composer = ProjectComposer(
            session,
            project_dir=os.path.join(HOME, "TIA"),
            project_name="ModularDemo",
        )

        composer.build(device_specs, io_card_specs, net)

        # Uncomment to import software:
        # composer.apply_software(sw_specs)

        hw_results = composer.compile_all_hw()
        sw_results = composer.compile_all_sw()

        log.info("HW compile results: %s", hw_results)
        log.info("SW compile results: %s", sw_results)

        composer.save()