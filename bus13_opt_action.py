# Kevin Moy, 8/6/2020
# Find optimal capacitor control on OpenDSS IEEE 13-bus given a load profile

import win32com.client
import pandas as pd
import os
import numpy as np
from bus13_state_reward import *

TOTAL_ACTIONS = 4


def cap_control(action, DSSCircuit):
    # Currently, only takes in IEEE 13 bus OpenDSS as input
    # action: Range of [0 3] of actions from RL agent
    # DSSCircuit: object of type DSSObj.ActiveCircuit (COM interface for OpenDSS Circuit)

    # Execute capacitor bank control based on action, given the following action space:

    if action == 0:
        # Both capacitors off:
        DSSCircuit.Capacitors.Name = "Cap1"
        DSSCircuit.Capacitors.States = (0,)
        DSSCircuit.Capacitors.Name = "Cap2"
        DSSCircuit.Capacitors.States = (0,)
    elif action == 1:
        # Capacitor 1 on, Capacitor 2 off:
        DSSCircuit.Capacitors.Name = "Cap1"
        DSSCircuit.Capacitors.States = (1,)
        DSSCircuit.Capacitors.Name = "Cap2"
        DSSCircuit.Capacitors.States = (0,)
    elif action == 2:
        # Capacitor 1 off, Capacitor 2 on:
        DSSCircuit.Capacitors.Name = "Cap1"
        DSSCircuit.Capacitors.States = (0,)
        DSSCircuit.Capacitors.Name = "Cap2"
        DSSCircuit.Capacitors.States = (1,)
    elif action == 3:
        # Both capacitors on:
        DSSCircuit.Capacitors.Name = "Cap1"
        DSSCircuit.Capacitors.States = (1,)
        DSSCircuit.Capacitors.Name = "Cap2"
        DSSCircuit.Capacitors.States = (1,)
    else:
        print("Invalid action " + str(action) + ", action in range [0 3] expected")


def opt_control(loadNames, loadKws, DSSCircuit, DSSSolution):
    for loadnum in range(np.size(loadNames)):
        DSSCircuit.SetActiveElement("Load." + loadNames[loadnum])
        # Set load with new loadKws
        DSSCircuit.ActiveDSSElement.Properties("kW").Val = loadKws[loadnum]

    # Get action with lowest reward
    opt_reward = -np.Inf
    opt_action = -np.Inf
    for action in range(TOTAL_ACTIONS):
        cap_control(action, DSSCircuit)
        DSSSolution.solve()
        observation = get_state(DSSCircuit)
        reward = quad_reward(observation)
        print("action=", action)
        print('reward=', reward)
        if reward > opt_reward:
            opt_action = action
            opt_reward = reward
    print("opt action = ", opt_action)
    print("opt reward = ", opt_reward)
    return opt_action, opt_reward
