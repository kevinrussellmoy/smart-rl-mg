import numpy as np
import random

import gym
import os

from stable_baselines.common.env_checker import check_env
from stable_baselines import DQN
from stable_baselines.common.evaluation import evaluate_policy

currentDirectory = os.getcwd()

env = gym.make('gym_openDSS:openDSS-v0')

check_env(env)

# Instantiate the agent
model = DQN('MlpPolicy', env, learning_rate=1e-3, prioritized_replay=True, verbose=1)
# Train the agent
model.learn(total_timesteps=int(1e3))
# Save the agent
model.save(currentDirectory + "\\opendss")
del model  # delete trained model to demonstrate loading

# Load the trained agent
model = DQN.load(currentDirectory + "\\opendss", env)

# Evaluate the agent
mean_reward, std_reward = evaluate_policy(model, model.get_env(), n_eval_episodes=10)

# Enjoy trained agent
obs = env.reset()
for i in range(1000):
    action, _states = model.predict(obs)
    obs, rewards, dones, info = env.step(action)
    env.render()



