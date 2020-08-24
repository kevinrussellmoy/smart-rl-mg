import pandas
import gym
import numpy as np
from sklearn.preprocessing import LabelEncoder
from keras.utils import np_utils
import tensorflow as tf
from tensorflow import keras

import os
currentDirectory = os.getcwd().replace('\\', '/')
datasetDirectory = currentDirectory.replace('/supervised-learning', '')

dataframe = pandas.read_csv(datasetDirectory + '/training_set.csv')
dataset = dataframe.T.values

voltages = dataset[1:,:41].astype(float)

action = dataset[1:,41]

action_onehot = np_utils.to_categorical(action) # Create one-hot encoding of action outputs

env = gym.make('gym_openDSS:openDSS-v0')
print('Number of actions: {}'.format(env.action_space.n))

# model = keras.Sequential([keras.layers.Flatten(input_shape=(1,) + env.observation_space.shape),
#                          keras.layers.Dense(len(env.VoltageMag), activation='relu'),
#                          keras.layers.Dense(len(env.VoltageMag), activation='relu'),
#                          keras.layers.Dense(len(env.VoltageMag), activation='relu'),
#                          keras.layers.Dense(4, activation='softmax')])

model = keras.Sequential([keras.layers.Flatten(input_shape=(1,) + env.observation_space.shape),
                         keras.layers.Dense(len(env.VoltageMag), activation='relu'),
                         keras.layers.Dense(len(env.VoltageMag), activation='relu'),
                         keras.layers.Dense(len(env.VoltageMag), activation='relu'),
                         keras.layers.Dense(4)])

# model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])
model.compile(optimizer='adam', loss=tf.keras.losses.SparseCategoricalCrossentropy(from_logits=True),
              metrics=['accuracy'])

# model.fit(voltages, action_onehot, epochs=100)
model.fit(voltages, action, epochs=100)

test_loss = test_acc = model.evaluate(voltages, action, verbose=2)