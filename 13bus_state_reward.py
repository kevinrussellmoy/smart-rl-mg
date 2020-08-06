# Kevin Moy, 8/6/2020
# Methods to read OpenDSS IEEE 13-bus and determine system state and reward from system states

import win32com.client
import numpy as np

# Upper and lower bounds of voltage zones:
ZONE2_UB = 1.10
ZONE1_UB = 1.05
ZONE1_LB = 0.95
ZONE2_LB = 0.90

# Penalties for each zone:
# TODO: Tune these hyperparameters
ZONE1_PENALTY = -200
ZONE2_PENALTY = -400


def get_state(DSSCircuit):
    # Currently, only takes in IEEE 13 bus OpenDSS as input
    # DSSCircuit: object of type DSSObj.ActiveCircuit (COM interface for OpenDSS Circuit)
    # Returns: NumPy array of voltages for each bus as the state vector

    vmags = DSSCircuit.AllBusVmagPu
    states = np.array(vmags)

    return states


def calc_reward(sts):
    # Calculate reward from states (bus voltages)
    # sts: NumPy array of voltages for each bus as the state vector
    # Returns: single floating-point number as reward

    # Number of buses in voltage zone 1
    num_zone1 = np.size(np.nonzero(np.logical_and(sts >= ZONE1_UB, sts < ZONE2_UB)))\
                + np.size(np.nonzero(np.logical_and(sts <= ZONE1_LB, sts > ZONE2_LB)))

    # Number of buses in voltage zone 2
    num_zone2 = np.size(np.nonzero(sts >= ZONE2_UB)) \
                + np.size(np.nonzero(sts <= ZONE2_LB))

    reward = num_zone1*ZONE1_PENALTY + num_zone2*ZONE2_PENALTY

    return reward
