import numpy as np
import random
import gym
import os
import matplotlib.pyplot as plt
import tensorflow as tf

from stable_baselines.common.env_checker import check_env
from stable_baselines import DQN
from stable_baselines.common.evaluation import evaluate_policy
from stable_baselines import results_plotter
from stable_baselines.bench import Monitor
from stable_baselines.results_plotter import load_results, ts2xy
from stable_baselines.common.noise import AdaptiveParamNoiseSpec
from stable_baselines.common.callbacks import BaseCallback

currentDirectory = os.getcwd()

# Set tensorflow print level
tf.logging.set_verbosity(tf.logging.ERROR)

class SaveOnBestTrainingRewardCallback(BaseCallback):
    """
    Callback for saving a model (the check is done every ``check_freq`` steps)
    based on the training reward (in practice, we recommend using ``EvalCallback``).

    :param check_freq: (int)
    :param log_dir: (str) Path to the folder where the model will be saved.
      It must contains the file created by the ``Monitor`` wrapper.
    :param verbose: (int)
    """
    def __init__(self, check_freq: int, log_dir: str, verbose=1):
        super(SaveOnBestTrainingRewardCallback, self).__init__(verbose)
        self.check_freq = check_freq
        self.log_dir = log_dir
        self.save_path = os.path.join(log_dir, 'best_model')
        self.best_mean_reward = -np.inf

    def _init_callback(self) -> None:
        # Create folder if needed
        if self.save_path is not None:
            os.makedirs(self.save_path, exist_ok=True)

    def _on_step(self) -> bool:
        if self.n_calls % self.check_freq == 0:

            # Retrieve training reward
            x, y = ts2xy(load_results(self.log_dir), 'timesteps')
            if len(x) > 0:
                # Mean training reward over the last 100 episodes
                mean_reward = np.mean(y[-100:])
                if self.verbose > 0:
                    print("Num timesteps: {}".format(self.num_timesteps))
                    print("Best mean reward: {:.2f} - Last mean reward per episode: {:.2f}".format(self.best_mean_reward, mean_reward))

                # New best model, you could save the agent here
                if mean_reward > self.best_mean_reward:
                    self.best_mean_reward = mean_reward
                    # Example for saving best model
                    if self.verbose > 0:
                        print("Saving new best model to {}".format(self.save_path))
                    self.model.save(self.save_path)
        return True


def moving_average(values, window):
    """
    Smooth values by doing a moving average
    :param values: (numpy array)
    :param window: (int)
    :return: (numpy array)
    """
    weights = np.repeat(1.0, window) / window
    return np.convolve(values, weights, 'valid')


def plot_results(log_folder, title='Learning Curve'):
    """
    plot the results

    :param log_folder: (str) the save location of the results to plot
    :param title: (str) the title of the task to plot
    """
    x, y = ts2xy(load_results(log_folder), 'timesteps')
    y = moving_average(y, window=50)
    # Truncate x
    x = x[len(x) - len(y):]

    fig = plt.figure(title)
    plt.plot(x, y)
    plt.xlabel('Number of Timesteps')
    plt.ylabel('Rewards')
    plt.title(title + " Smoothed")
    plt.show()

# Create and wrap the environment
env = gym.make('gym_openDSS:openDSS-v0')
check_env(env)  # Check env to make sure it conforms to gym requirements
env = Monitor(env, currentDirectory)

# Instantiate the agent
model = DQN('MlpPolicy', env, learning_rate=1e-3, prioritized_replay=True, verbose=1)

# Create the callback: check every 100 steps
callback = SaveOnBestTrainingRewardCallback(check_freq=100, log_dir=currentDirectory)

# Train the agent
time_steps = 1e4
model.learn(total_timesteps=int(time_steps))
# Save the agent
model.save(currentDirectory + "\\opendss_1e4")

results_plotter.plot_results([currentDirectory], time_steps, results_plotter.X_TIMESTEPS, "DQN IEEE 13-bus")
plt.show()

# del model  # delete trained model to demonstrate loading

# Load the trained agent
model = DQN.load(currentDirectory + "\\opendss_1e4", env)

# Evaluate the agent
mean_reward, std_reward = evaluate_policy(model, model.get_env(), n_eval_episodes=10)

print("Mean reward, ", mean_reward, " std reward, ", std_reward)
#
# # Enjoy trained agent
# obs = env.reset()
# for i in range(1000):
#     action, _states = model.predict(obs)
#     obs, rewards, dones, info = env.step(action)



