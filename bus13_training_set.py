# produce CSV of voltages (pre-action) and known optimal action for later supervised learning

# TODO: Reincorporate into bus13_load_configs?

from bus13_opt_action import *
from bus13_state_reward import *
from bus13_capcontrol_action import *

import numpy as np
import os

currentDirectory = os.getcwd().replace('\\', '/')

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

# Instantiate CapControl elements
DSSText.Command = "New CapControl.CtrlCap1 Capacitor=Cap1 element=Line.684611 terminal=1 type=Voltage ptratio=2400"
DSSText.Command = "~ ctratio=1 ptphase=MIN ONsetting=0.90 OFFsetting=1.10 Delay=50 Delayoff=100"
DSSText.Command = "Disable CapControl.CtrlCap1"

DSSText.Command = "New CapControl.CtrlCap2 Capacitor=Cap2 element=Line.692675 terminal=1 type=Voltage ptratio=2400"
DSSText.Command = "~ ctratio=1 ptphase=MIN ONsetting=0.90 OFFsetting=1.10 Delay=50 Delayoff=100"
DSSText.Command = "Disable CapControl.CtrlCap2"

loadNames = np.array(DSSCircuit.Loads.AllNames)

# Initially disable both capacitor banks and set both to 1500 KVAR rating
capNames = DSSCircuit.Capacitors.AllNames
for cap in capNames:
    DSSCircuit.SetActiveElement("Capacitor." + cap)
    DSSCircuit.ActiveDSSElement.Properties("kVAR").Val = 1500
DSSCircuit.Capacitors.Name = "Cap1"
DSSCircuit.Capacitors.States = (0,)
DSSCircuit.Capacitors.Name = "Cap2"
DSSCircuit.Capacitors.States = (0,)

# Load in load configurations from testing database
test_load_kw = pd.read_csv(currentDirectory + "/loadkW_train.csv") # generate this from bus13_load_configs
traindf = pd.DataFrame()

for i in np.arange(1,10000):
    loadKws = np.array(test_load_kw[str(i)])

    for loadnum in range(np.size(loadNames)):
        DSSCircuit.SetActiveElement("Load." + loadNames[loadnum])
        # Set load with new loadKws
        DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]
    # Reset capacitors in between each test
    action_to_cap_control(0, DSSCircuit)
    observation = get_state(DSSCircuit)

    # ****************************************************
    # * Known Optimal Action
    # ****************************************************
    # Computes action and reward for each action and chooses action with highest reward
    opt_action, opt_reward = opt_control(DSSCircuit, DSSSolution)
    obsAndAction = np.append(observation, opt_action)
    traindf[i] = obsAndAction


traindf.to_csv(currentDirectory + "/training_set.csv")
