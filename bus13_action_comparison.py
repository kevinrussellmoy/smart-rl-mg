# Comparison between RL control, optimal (known least penalty) control, and OpenDSS CapControl

import win32com.client
import pandas as pd
import os
import numpy as np
from bus13_opt_action import *
from bus13_state_reward import *

currentDirectory = os.getcwd()

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

loadNames = np.array(DSSCircuit.Loads.AllNames)

# Load in load configuration from testing database
# TODO: Run all of this in a loop across all load configs in loadkW_test
test_load_kw = pd.read_csv(currentDirectory + "\\loadkW_test.csv")
loadKws = np.array(test_load_kw[str(1)])

# Reset capacitors in between each test
cap_control(0, DSSCircuit)
observation = get_state(DSSCircuit)

# ****************************************************
# * Trained RL Agent Control
# ****************************************************
# TODO: Load in trained model; predict action from observation, with output of rl_action and rl_reward


# Reset capacitors in between each test
cap_control(0, DSSCircuit)
# ****************************************************
# * Known Optimal Action
# ****************************************************
# Computes action and reward for each action and chooses action with highest reward
opt_action, opt_reward = opt_control(loadNames, loadKws, DSSCircuit, DSSSolution)


# Reset capacitors in between each test
cap_control(0, DSSCircuit)
# ****************************************************
# * OpenDSS CapControl ACtion
# ****************************************************
# TODO: Action based on CapControl object in OpenDSS as configured in XXX.py

