import numpy as np
import random

import gym

import sys

from tensorflow.keras import Model, Sequential
from tensorflow.keras.layers import Dense, Flatten, Reshape, Activation
from tensorflow.keras.optimizers import Adam

from rl.agents.dqn import DQNAgent
from rl.policy import EpsGreedyQPolicy
from rl.memory import SequentialMemory

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
sys.stdout = open("C:/Users/chris/OneDrive/Desktop/Miscellaneous/Bits&Watts/test4.txt", "w")
dqn.fit(env, nb_steps=1000, visualize=True, verbose=2)
sys.stdout.close()


