# Comparison between RL control, optimal (known least penalty) control, and OpenDSS CapControl

from bus13_opt_action import *
from bus13_state_reward import *

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

# For storing action/reward pairs
labels = [np.array(['agent', 'agent', 'opt', 'opt', 'capctrl', 'capctrl']), np.array(['action', 'reward', 'action', 'reward', 'action', 'reward'])]
act_reward_array = np.zeros((6, 1000))

# Load in load configurations from testing database
test_load_kw = pd.read_csv(currentDirectory + "\\loadkW_test.csv")

for i in np.arange(1,10):
    loadKws = np.array(test_load_kw[str(i)])

    for loadnum in range(np.size(loadNames)):
        DSSCircuit.SetActiveElement("Load." + loadNames[loadnum])
        # Set load with new loadKws
        DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]

    # Reset capacitors in between each test
    cap_control(0, DSSCircuit)
    observation = get_state(DSSCircuit)

    # ****************************************************
    # * Trained RL Agent Control
    # ****************************************************

    # load model
    model = load_model(currentDirectory + "/keras-rl/saved_model")
    # Configure DQN agent, including weights
    memory = SequentialMemory(limit=50000, window_length=1)
    dqn = DQNAgent(model=model, nb_actions=4, memory=memory)
    print("Loaded model from disk")

    obs_array = observation.reshape(1, 1, len(observation))
    print(dqn.model.predict(obs_array))
    rl_action = np.argmax(dqn.model.predict(obs_array))
    # action = 3
    print("rl action: ", rl_action)
    cap_control(rl_action, DSSCircuit)
    DSSSolution.solve()
    rl_reward = quad_reward(get_state(DSSCircuit))
    print("rl reward: ", quad_reward(rl_reward))
    print("\n")

    act_reward_array[(0, i)] = rl_action
    act_reward_array[(1, i)] = rl_reward

    # Reset capacitors in between each test
    cap_control(0, DSSCircuit)

    # ****************************************************
    # * Known Optimal Action
    # ****************************************************
    # Computes action and reward for each action and chooses action with highest reward
    opt_action, opt_reward = opt_control(DSSCircuit, DSSSolution)

    act_reward_array[(2, i)] = opt_action
    act_reward_array[(3, i)] = opt_reward

    # Reset capacitors in between each test
    cap_control(0, DSSCircuit)

    # ****************************************************
    # * OpenDSS CapControl Action
    # ****************************************************
    # TODO: Action based on CapControl object in OpenDSS as configured in XXX.py
    capctrl_action = 0
    capctrl_reward = -200

    act_reward_array[(4, i)] = capctrl_action
    act_reward_array[(5, i)] = -200

act_reward_df = pd.DataFrame(act_reward_array, index = labels)
act_reward_df.to_csv(currentDirectory + "/action_reward_comparison.csv")
# TODO: Capture all action/reward pairs in Dataframe
