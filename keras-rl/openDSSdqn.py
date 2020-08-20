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
from keras.models import model_from_json
from matplotlib.figure import Figure

from rl.agents.dqn import DQNAgent
from rl.policy import EpsGreedyQPolicy, LinearAnnealedPolicy
from rl.memory import SequentialMemory

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
    loss = np.array(history.history['episode_reward'])
    epochs = np.arange(loss.size)

    # fig = plt.figure(figsize=(10, 10))
    # plt.xlabel("Epochs")
    # plt.ylabel("Loss")
    # TODO: Clean this mess up lol
    window = 100
    title = 'Learning Curve'
    weights = np.repeat(1.0, window) / window
    y = np.convolve(loss, weights, 'valid')
    x = epochs
    plt.plot(loss)

    # Truncate x
    x = x[len(x) - len(y):]
    fig = plt.figure(title)
    plt.plot(x, y)
    plt.xlabel('Number of Timesteps')
    plt.ylabel('Rewards')
    plt.title(title + " Smoothed")

    if output_path:
        os.makedirs(output_path, exist_ok=True)
        plt.savefig(os.path.join(output_path, 'loss_epochs_smoothed.%s' % extension))

    plt.show()

    return fig


# Instantiate environment
env = gym.make('gym_openDSS:openDSS-v0')
print('Number of actions: {}'.format(env.action_space.n))

model = Sequential()
model.add(Flatten(input_shape=(1,) + env.observation_space.shape))
model.add(Dense(len(env.VoltageMag)))
model.add(Activation('relu'))
# TODO: Add additional layers here
model.add(Dense(len(env.VoltageMag)))
model.add(Activation('relu'))
model.add(Dense(env.action_space.n))
model.add(Activation('linear'))
print(model.summary())
#
policy = LinearAnnealedPolicy(EpsGreedyQPolicy(), attr='eps', value_max=1, value_min=.2, value_test=.05, nb_steps=1000)
memory = SequentialMemory(limit=50000, window_length=1)
dqn = DQNAgent(model=model, nb_actions=env.action_space.n, memory=memory, nb_steps_warmup=1000,
               target_model_update=1e-3, policy=policy)
dqn.compile(Adam(lr=1e-1), metrics=['mae'])
#
# sys.stdout = open(currentDirectory + "/test4.txt", "w")
history = dqn.fit(env, nb_steps=10000, verbose=2)
# sys.stdout.close()

# Save model
dqn.model.save(currentDirectory + "/saved_model")

print("Saved model to disk")

# Enjoy trained agent
for i in range(50):
    obs = env.reset()
    obs_array = obs.reshape(1,1,len(env.VoltageMag))
    print(dqn.model.predict(obs_array))
    action = np.argmax(dqn.model.predict(obs_array))
    print("action: ", action)
    observs, rewards, dones, info = env.step(action)
    print("reward: ", rewards)


fig = plot_from_history(history, currentDirectory)