# Kevin Moy, 7/24/2020
# Uses IEEE 13-bus system to scale loads and determine state space for RL agent learning
# Return vector of load values given load inputs

import win32com.client
import pandas as pd
import os
import numpy as np

MAX_NUM_CONFIG = 20
MIN_BUS_VOLT = 0.8
MAX_BUS_VOLT = 1.2


def load_states(loadNames, DSSCircuit, DSSSolution, min_load=0.5, max_load=3):
    # Currently, only takes in IEEE 13 bus OpenDSS as input
    # loadNames: vector of bus names for each load
    # DSSCircuit: object of type DSSObj.ActiveCircuit (COM interface for OpenDSS Circuit)
    # DSSSolution: object of type DSSObj.ActiveCircuit.Solution

    # Randomly scales loads with uniform distribution between min_load and max_load (percentages)
    # Then solves circuit with all capacitors off, then all capacitors on
    # If the voltages fall within the MIN_BUS_VOLT and MAX_BUS_VOLT range,
    # then the load configuration is returned as an array

    # Array of random numbers to scale loads by:
    randScale = np.random.uniform(min_load, max_load, loadNames.__len__())

    scale_up(DSSCircuit, randScale)

    # Initially disable both capacitor banks and set both to 1500 KVAR rating
    capNames = DSSCircuit.Capacitors.AllNames
    for cap in capNames:
        DSSCircuit.SetActiveElement("Capacitor." + cap)
        DSSCircuit.ActiveDSSElement.Properties("kVAR").Val = 1500

    DSSCircuit.Capacitors.Name = "Cap1"
    DSSCircuit.Capacitors.States = (0,)
    DSSCircuit.Capacitors.Name = "Cap2"
    DSSCircuit.Capacitors.States = (0,)

    # Solve the Circuit
    DSSSolution.Solve()
    # if DSSSolution.Converged:
    #     print("The Circuit Solved Successfully")
    # else:
    #     print("The Circuit Did Not Solve Successfully")

    # Retrieve all voltage magnitudes from all phases of all buses
    VoltageMagCapOff = DSSCircuit.AllBusVmagPu

    maxVCapOff = max(VoltageMagCapOff)
    minVCapOff = min(VoltageMagCapOff)

    # print("Maximum Voltage is: " + str(maxVCapOff))
    # print("Minimum Voltage is: " + str(minVCapOff))

    # ----- DETERMINING STATE SPACE -----
    # print("Enabling Capacitor Banks")
    # Enable both capacitor banks
    DSSCircuit.Capacitors.Name = "Cap1"
    DSSCircuit.Capacitors.States = (1,)
    DSSCircuit.Capacitors.Name = "Cap2"
    DSSCircuit.Capacitors.States = (1,)

    # Solve the Circuit
    DSSSolution.Solve()
    # if DSSSolution.Converged:
    #     print("The Circuit Solved Successfully")
    # else:
    #     print("The Circuit Did Not Solve Successfully")

    # Retrieve all voltage magnitudes from all phases of all buses
    VoltageMagCapOn = DSSCircuit.AllBusVmagPu

    maxVCapOn = max(VoltageMagCapOn)
    minVCapOn = min(VoltageMagCapOn)

    # print("Maximum Voltage is: " + str(maxVCapOn))
    # print("Minimum Voltage is: " + str(minVCapOn))

    if (max(maxVCapOff, maxVCapOn, minVCapOff, minVCapOn) <= MAX_BUS_VOLT) & \
            (min(maxVCapOff, maxVCapOn, minVCapOff, minVCapOn) >= MIN_BUS_VOLT):
        print("Voltages within acceptable range")
        loadKws = []
        for load in loadNames:
            DSSCircuit.SetActiveElement("Load." + load)
            # Get bus of load
            kwLoad = DSSCircuit.ActiveDSSElement.Properties("kW").Val
            loadKws.append(kwLoad)
        scale_down(DSSCircuit, randScale)
        return np.array(loadKws)
    else:
        print("Voltages not within acceptable range [" + str(MIN_BUS_VOLT) + ", " + str(MAX_BUS_VOLT) +
              "] p.u., not saving")
        scale_down(DSSCircuit, randScale)
        return None


def scale_up(DSSCircuit, randScale):
    # Step through every load and scale it up by a random percentage
    iLoads = DSSCircuit.Loads.First
    rdx = 0
    while iLoads:
        # Scale load by 10000%
        DSSCircuit.Loads.kW = DSSCircuit.Loads.kW * randScale[rdx]
        # Move to next load and random number
        iLoads = DSSCircuit.Loads.Next
        rdx += 1


def scale_down(DSSCircuit, randScale):
    # Step through every load and scale back down by the same random percentage
    iLoads = DSSCircuit.Loads.First
    rdx = 0
    while iLoads:
        # Scale load by 10000%
        DSSCircuit.Loads.kW = DSSCircuit.Loads.kW / randScale[rdx]
        # Move to next load and random number
        iLoads = DSSCircuit.Loads.Next
        rdx += 1
