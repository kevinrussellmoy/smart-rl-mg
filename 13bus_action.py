# Kevin Moy, 8/6/2020
# Method to read in action from reinforcement agent and
# execute capacitor control on OpenDSS IEEE 13-bus

import win32com.client


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


