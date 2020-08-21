# Kevin Moy, 8/6/2020
# Find capacitor control from OpenDSS CapControl control object, on OpenDSS IEEE 13-bus, given a load profile

import win32com.client
import pandas as pd
import os
import numpy as np
from bus13_opt_action import *

TOTAL_ACTIONS = 4

# TODO: Write params and returns; add cap kVARs as params


def cap_control(DSSCircuit, DSSSolution, DSSText):
    # Initial solve for voltage states for cap control
    DSSSolution.Solve()

    # Enable cap control
    DSSText.Command = "Enable CapControl.CtrlCap1"
    DSSText.Command = "Enable CapControl.CtrlCap2"

    # Solve for cap control given initial solve voltage states
    DSSSolution.Solve()

    DSSCircuit.Capacitors.Name = "Cap1"
    cap1_state = DSSCircuit.Capacitors.States
    DSSCircuit.Capacitors.Name = "Cap2"
    cap2_state = DSSCircuit.Capacitors.States

    # TODO: damn I wish I had knew how to use dictionaries
    capctrl_action = -1000
    if (cap1_state[0] == 0) & (cap2_state[0] == 0):
        capctrl_action = 0
    elif (cap1_state[0] == 1) & (cap2_state[0] == 0):
        capctrl_action = 1
    elif (cap1_state[0] == 0) & (cap2_state[0] == 1):
        capctrl_action = 2
    elif (cap1_state[0] == 1) & (cap2_state[0] == 1):
        capctrl_action = 3
    else:
        print("Invalid action, action in range [0 3] expected")

    # Reset cap control (for some reason, CtrlCap1 adds a bus.... not sure why :(
    DSSText.Command = "Disable CapControl.CtrlCap1"
    DSSText.Command = "Disable CapControl.CtrlCap2"

    # Now solve with manual capacitor control
    action_to_cap_control(capctrl_action, DSSCircuit)
    DSSSolution.Solve()
    capctrl_reward = quad_reward(get_state(DSSCircuit))

    print("\nCapControl action = ", capctrl_action)
    print("CapControl reward = ", capctrl_reward)

    return capctrl_action, capctrl_reward
