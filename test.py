from tia_composer.tia_composer import TiaSession, DeviceSpec, NetworkSpec, ProjectComposer
from tia_composer.tia_hmi_manager import get_hmi_manager_from_composer, HmiScreenSpec, ScreenItemSpec
from pathlib import Path
import os
import shutil

specs = [
    DeviceSpec("PLC1",    "OrderNumber:6ES7 513-1AL02-0AB0/V2.6",  None, ip="192.168.0.130"),
    DeviceSpec("IOnode1", "OrderNumber:6ES7 155-6AU01-0BN0/V4.1",  None, ip="192.168.0.131"),
    DeviceSpec("HMI1",    "OrderNumber:6AV2 124-0MC01-0AX0/17.0.0.0", None, ip="192.168.0.132"),
]
io_cards = [
    DeviceSpec("IO1", "OrderNumber:6ES7 521-1BL00-0AB0/V2.1", slot=2, parent="PLC1"),
    DeviceSpec("IO1", "OrderNumber:6ES7 131-6BH01-0BA0/V0.0", slot=1, parent="IOnode1"),
]
net = NetworkSpec(subnet_name="Profinet", io_system_name="PNIO")

project_dir = r"C:\TIA"
project_name = "ModularDemo"

project_path = Path(project_dir)/Path(project_name)
print(project_path,project_path.exists())
if project_path.exists():
    shutil.rmtree(project_path)

with TiaSession(ui=False) as session:

    composer = ProjectComposer(session, project_dir=project_dir, project_name=project_name)
    composer.build(specs, io_cards, net)
    hmi = get_hmi_manager_from_composer(composer, "HMI1")
    # hmi.create(HmiScreenSpec("MainScreen", is_start_screen=True, items=[
    #     ScreenItemSpec("Label", "Title", left=40, top=20, width=400, height=50,
    #                 attributes={"Text": "Line 1"}),
    # ]))
    composer.compile_all_hw()
    composer.compile_all_sw()
    composer.save()