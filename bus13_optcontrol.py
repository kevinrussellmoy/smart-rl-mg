# Kevin Moy 8/12/2020
# Method to instantiate OpenDSS IEEE 13-bus and find optimal action (control of capacitors) based on lowest reward

# TODO: this

import win32com.client
import numpy as np

NUM_STEPS = 4

def find_opt_action(loadkws):
    # input: loadkws: NumPy array of load configurations
    # output: opt_action, opt_reward: optimal action, optimal reward tuple

    # Instantiate environment for testing
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

    # Test all four actions
    opt_reward = -np.Inf
    opt_action = -np.Inf

    for step in range(NUM_STEPS):
        action = step
        obs, reward, done, info = env.step(action)
        if reward > opt_reward:
            opt_reward = reward
            opt_action = action

    return opt_action, opt_reward
