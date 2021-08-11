import matplotlib.pyplot as plt
import numpy as np
from numpy import random

class SingleNeuron(object):
    def __init__(self):
        self._w = 0.000
        self._b = 0.000
        self._x = 0.000

    def set_params(self, w, b):
        self._w = w
        self._b = b

    def for_pass(self, x):
        self._x = x
        _y_hat = self._w * self._x + self._b
        return _y_hat

    def back_prop(self, err, cost=0.1):
        m = len(self._x)
        self._w_grad = cost * np.sum(err * self._x) / m
        self._b_grad = cost * np.sum(err * 1) / m

    def update_grad(self):
        self.set_params(self._w + self._w_grad, self._b + self._b_grad)

    def fit(self, x, y, n_iter=100):
        for i in range(n_iter):
            y_hat = self.for_pass(x)
            error = y - y_hat
            self.back_prop(error)
            self.update_grad()


n1 = SingleNeuron()
n1.set_params(1, 0)

# Set plot groups
x = random.rand(1000)
y = random.rand(1000)

#for i in range(len(x)):
#    y[i] += x[i] * (-2.5)

# Draws plot groups
xpoints = np.array(x)
ypoints = np.array(y)
plt.scatter(x, y, alpha=0.5)

print(f"w:{n1._w:>8.5f} / b:{n1._b:>8.5f}")

# Calculate
crital_value= 1e-10
for i in range(100):
    n1.fit(x, y, 100)
    print(f"w:{n1._w:>8.5f} / b:{n1._b:>8.5f}")
    if (abs(n1._w_grad) <= crital_value and abs(n1._b_grad) <= crital_value):
        break 

print(f"w:{n1._w:>8.5f} / b:{n1._b:>8.5f}")

# Draws calculated result
line_x = np.array([np.min(x), np.max(x)])
line_y = n1.for_pass(line_x)
plt.plot(line_x, line_y, 'r')
plt.show()
