import numpy as np
import random

import gym
import os
import sys
import matplotlib.pyplot as plt
import logging
logging.getLogger("tensorflow").setLevel(logging.ERROR) # Set tensorflow print level

from tensorflow.keras import Model, Sequential
from tensorflow.keras.layers import Dense, Flatten, Reshape, Activation
from tensorflow.keras.optimizers import Adam
from keras.callbacks import History
from matplotlib.figure import Figure

from rl.agents.dqn import DQNAgent
from rl.policy import EpsGreedyQPolicy
from rl.memory import SequentialMemory

PLOT_LOSS_FORMAT = 'loss_epochs.%s'

currentDirectory = os.getcwd().replace('\\', '/')


def plot_from_history(history: History,
                      output_path: str = None,
                      extension: str = 'pdf') -> Figure:
    """
    From https://github.com/rodrigobressan/entity_embeddings_categorical/blob/master/entity_embeddings/util/visualization_utils.py
    Used to make a Figure object containing the loss curve between the epochs.
    :param history: the history outputted from the model.fit method
    :param output_path: (optional) where the image will be saved
    :param extension: (optional) the extension of the file
    :return: a Figure object containing the plot
    """
    loss = history.history['episode_reward']

    fig = plt.figure(figsize=(10, 10))
    plt.xlabel("Epochs")
    plt.ylabel("Loss")

    plt.plot(loss)

    if output_path:
        os.makedirs(output_path, exist_ok=True)
        plt.savefig(os.path.join(output_path, PLOT_LOSS_FORMAT % extension))

    plt.show()

    return fig


# Instantiate environment
env = gym.make('gym_openDSS:openDSS-v0')
print('Number of actions: {}'.format(env.action_space.n))

model = Sequential()
model.add(Flatten(input_shape=(1,) + env.observation_space.shape))
model.add(Dense(len(env.VoltageMag)))
model.add(Activation('relu'))
model.add(Dense(env.action_space.n))
model.add(Activation('linear'))
print(model.summary())
#
policy = EpsGreedyQPolicy()
memory = SequentialMemory(limit=50000, window_length=1)
dqn = DQNAgent(model=model, nb_actions=env.action_space.n, memory=memory, nb_steps_warmup=500,
               target_model_update=1e-3, policy=policy)
dqn.compile(Adam(lr=1e-4), metrics=['mae'])
#
sys.stdout = open(currentDirectory + "/test4.txt", "w")
history = dqn.fit(env, nb_steps=1000, verbose=2)
sys.stdout.close()

fig = plot_from_history(history, currentDirectory)