import gym

env = gym.make('gym_openDSS:openDSS-v0')

# Box(41,) means that it is a Vector with 41 components
print("Observation space:", env.observation_space)
# Discrete(4) means that there is two discrete actions
print("Action space:", env.action_space)

# Get initial set of observations (states)
obs = env.reset()

# Test actions
nsteps = 4
for step in range(nsteps):
    print("\nStep: ", step)
    # action = env.action_space.sample()
    action = step
    print("Sampled action:", action)
    obs, reward, done, info = env.step(action)
    print('obs=', obs)
    print('reward=', reward)
    print ('done=', done)

