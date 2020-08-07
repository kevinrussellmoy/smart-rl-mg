# Kevin Moy, 8/6/2020
# Uses IEEE 13-bus system
# Test to see if we can set loads using generate_state_space.py

import win32com.client
import pandas as pd
import numpy as np
from generate_state_space import load_states

# Instantiate the OpenDSS Object
try:
    DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
except:
    print("Unable to start the OpenDSS Engine")
    raise SystemExit
print("OpenDSS Engine started\n")

# Set up the Text, Circuit, and Solution Interfaces
DSSText = DSSObj.Text
DSSCircuit = DSSObj.ActiveCircuit
DSSSolution = DSSCircuit.Solution

# Load in an example circuit
DSSText.Command = r"Compile 'C:\Program Files\OpenDSS\IEEETestCases\13Bus\IEEE13Nodeckt.dss'"

# Disable voltage regulators
DSSText.Command = "Disable regcontrol.Reg1"
DSSText.Command = "Disable regcontrol.Reg2"
DSSText.Command = "Disable regcontrol.Reg3"

# Solve initial circuit
DSSSolution.solve()
print(DSSCircuit.AllBusVmagPu)

loadNames = np.array(DSSCircuit.Loads.AllNames)
loadKwdf = pd.DataFrame(loadNames)

loadKws = load_states(loadNames, DSSCircuit, DSSSolution)

for loadnum in range(np.size(loadNames)):
    DSSCircuit.SetActiveElement("Load." + loadNames[loadnum])
    # Set load with new loadKws
    DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]

# Solve new circuit
DSSSolution.solve()
print(DSSCircuit.AllBusVmagPu)