# Comparison between RL control, optimal (known least penalty) control, and OpenDSS CapControl

from bus13_opt_action import *
from bus13_state_reward import *
from bus13_capcontrol_action import *

import numpy as np

import os
import matplotlib.pyplot as plt
import logging

logging.getLogger("tensorflow").setLevel(logging.ERROR)  # Set tensorflow print level

from keras.models import load_model

from rl.agents.dqn import DQNAgent
from rl.memory import SequentialMemory

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

# Instantiate CapControl elements
DSSText.Command = "New CapControl.CtrlCap1 Capacitor=Cap1 element=Line.684611 terminal=1 type=Voltage ptratio=2400"
DSSText.Command = "~ ctratio=1 ptphase=MIN ONsetting=0.90 OFFsetting=1.10 Delay=50 Delayoff=100"
DSSText.Command = "Disable CapControl.CtrlCap1"

DSSText.Command = "New CapControl.CtrlCap2 Capacitor=Cap2 element=Line.692675 terminal=1 type=Voltage ptratio=2400"
DSSText.Command = "~ ctratio=1 ptphase=MIN ONsetting=0.90 OFFsetting=1.10 Delay=50 Delayoff=100"
DSSText.Command = "Disable CapControl.CtrlCap2"

# load DQN agent model
model = load_model(currentDirectory + "/keras-rl/saved_model")
# Configure DQN agent, including weights
memory = SequentialMemory(limit=50000, window_length=1)
dqn = DQNAgent(model=model, nb_actions=4, memory=memory)
print("Loaded model from disk")

# For storing action/reward pairs
labels = [np.array(['agent', 'agent', 'opt', 'opt', 'capctrl', 'capctrl']),
          np.array(['action', 'reward', 'action', 'reward', 'action', 'reward'])]
act_reward_array = np.zeros((6, 1000))

# Load in load configurations from testing database
test_load_kw = pd.read_csv(currentDirectory + "\\loadkW_test.csv")

for i in np.arange(1, 1000):
    loadKws = np.array(test_load_kw[str(i)])

    for loadnum in range(np.size(loadNames)):
        DSSCircuit.SetActiveElement("Load." + loadNames[loadnum])
        # Set load with new loadKws
        DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]

    # Reset capacitors in between each test
    action_to_cap_control(0, DSSCircuit)
    observation = get_state(DSSCircuit)
    print("\n")
    # ****************************************************
    # * Trained RL Agent Control
    # ****************************************************

    obs_array = observation.reshape(1, 1, len(observation))
    print(dqn.model.predict(obs_array))
    rl_action = np.argmax(dqn.model.predict(obs_array))
    # action = 3
    print("rl action: ", rl_action)
    action_to_cap_control(rl_action, DSSCircuit)
    DSSSolution.solve()
    rl_reward = quad_reward(get_state(DSSCircuit))
    print("rl reward: ", rl_reward)

    act_reward_array[(0, i)] = rl_action
    act_reward_array[(1, i)] = rl_reward

    # Reset capacitors in between each test
    action_to_cap_control(0, DSSCircuit)

    # ****************************************************
    # * Known Optimal Action
    # ****************************************************
    # Computes action and reward for each action and chooses action with highest reward
    opt_action, opt_reward = opt_control(DSSCircuit, DSSSolution)

    act_reward_array[(2, i)] = opt_action
    act_reward_array[(3, i)] = opt_reward

    # Reset capacitors in between each test
    action_to_cap_control(0, DSSCircuit)

    # ****************************************************
    # * OpenDSS CapControl Action
    # ****************************************************
    capctrl_action, capctrl_reward = cap_control(DSSCircuit, DSSSolution, DSSText)

    # # Reset capacitors in between each test
    # action_to_cap_control(0, DSSCircuit)
    #
    # action_to_cap_control(capctrl_action, DSSCircuit)
    # DSSSolution.solve()
    # capctrl_reward = quad_reward(get_state(DSSCircuit))

    act_reward_array[(4, i)] = capctrl_action
    act_reward_array[(5, i)] = capctrl_reward

act_reward_df = pd.DataFrame(act_reward_array, index=labels)
act_reward_df.to_csv(currentDirectory + "/action_reward_comparison.csv")
