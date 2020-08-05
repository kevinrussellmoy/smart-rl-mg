# Kevin Moy, 7/24/2020
# Uses IEEE 13-bus system to scale loads and determine state space for RL agent learning
# Produces CSV of load names and load magnitudes in kW for later testing

import win32com.client
import pandas as pd
import os
import numpy as np

currentDirectory = os.getcwd()  # Will switch to OpenDSS directory after OpenDSS Object is instantiated!

# Instantiate the OpenDSS Object
try:
    DSSObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
except:
    print("Unable to start the OpenDSS Engine")
    raise SystemExit
print("OpenDSS Engine started")

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

loadKwdf = pd.DataFrame(DSSCircuit.Loads.AllNames)
for i in range(5):
    print("Configuration " + str(i))
    # Get all load names
    loadNames = DSSCircuit.Loads.AllNames

    # Step through every load and scale it up by a random percentage between 50% and 200%
    iLoads = DSSCircuit.Loads.First
    while iLoads:
        # Scale load by 10000%
        DSSCircuit.Loads.kW = DSSCircuit.Loads.kW * np.random.uniform(0.5,2)
        # Move to next load
        iLoads = DSSCircuit.Loads.Next

    # Initially disable both capacitor banks and set both to 1500 KVAR rating
    capNames = DSSCircuit.Capacitors.AllNames
    for cap in capNames:
        DSSCircuit.SetActiveElement("Capacitor." + cap)
        DSSCircuit.ActiveDSSElement.Properties("kVAR").Val = 1500  # try changing to 1500, 2000
        # DSSCircuit.Disable(cap)

    DSSCircuit.Capacitors.Name = "Cap1"
    DSSCircuit.Capacitors.States = (0,)
    DSSCircuit.Capacitors.Name = "Cap2"
    DSSCircuit.Capacitors.States = (0,)

    # Solve the Circuit
    DSSSolution.Solve()
    if DSSSolution.Converged:
        print("The Circuit Solved Successfully")
    else:
        print("The Circuit Did Not Solve Successfully")

    # Retrieve all voltage magnitudes from all phases of all buses
    VoltageMagCapOff = DSSCircuit.AllBusVmagPu

    maxVCapOff = max(VoltageMagCapOff)
    minVCapOff = min(VoltageMagCapOff)

    print("Maximum Voltage is: " + str(maxVCapOff))
    print("Minimum Voltage is: " + str(minVCapOff))

    # ----- DETERMINING STATE SPACE -----
    print("Enabling Capacitor Banks")
    # Enable both capacitor banks
    DSSCircuit.Capacitors.Name = "Cap1"
    DSSCircuit.Capacitors.States = (1,)
    DSSCircuit.Capacitors.Name = "Cap2"
    DSSCircuit.Capacitors.States = (1,)

    # Solve the Circuit
    DSSSolution.Solve()
    if DSSSolution.Converged:
        print("The Circuit Solved Successfully")
    else:
        print("The Circuit Did Not Solve Successfully")

    # Retrieve all voltage magnitudes from all phases of all buses
    VoltageMagCapOn = DSSCircuit.AllBusVmagPu

    maxVCapOn = max(VoltageMagCapOn)
    minVCapOn = min(VoltageMagCapOn)

    print("Maximum Voltage is: " + str(maxVCapOn))
    print("Minimum Voltage is: " + str(minVCapOn))

    if (max(maxVCapOff, maxVCapOn, minVCapOff, minVCapOn) <= 1.25) & (min(maxVCapOff, maxVCapOn, minVCapOff, minVCapOn) >= 0.80):
        print("Voltages within acceptable range")
        loadKws = []
        for load in loadNames:
            DSSCircuit.SetActiveElement("Load." + load)
            # Get bus of load
            kwLoad = DSSCircuit.ActiveDSSElement.Properties("kW").Val
            loadKws.append(kwLoad)
        loadKwdf[i+1] = loadKws
    else:
        print("Voltages not within acceptable range [0.85, 1.15] p.u., not saving")

loadKwdf.set_index(0, inplace=True)
loadKwdf.to_csv(currentDirectory + "\\loadkW.csv", index_label='Load Name')
print("saved to "+ currentDirectory + "\\loadkW.csv")

