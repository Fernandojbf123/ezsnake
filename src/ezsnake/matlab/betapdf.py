import numpy as np
from scipy.special import betaln


def betapdf(x, a, b):
    """
    BETAPDF Beta probability density function.
    Y = BETAPDF(X,A,B) returns the beta probability density
    function with parameters A and B at the values in X.

    The size of Y is the common size of the input arguments. A scalar input
    functions as a constant matrix of the same size as the other inputs.

    See also BETACDF, BETAFIT, BETAINV, BETALIKE, BETARND, BETASTAT, PDF,
             BETA, BETALN.

    References:
       [1]  M. Abramowitz and I. A. Stegun, "Handbook of Mathematical
       Functions", Government Printing Office, 1964, 26.1.33.
    """
    x = np.asarray(x)
    a = np.asarray(a)
    b = np.asarray(b)
    # Broadcast to common shape
    x, a, b = np.broadcast_arrays(x, a, b)
    y = np.zeros_like(x, dtype=float)
    # Special cases
    y[(a == 1) & (x == 0)] = b[(a == 1) & (x == 0)]
    y[(b == 1) & (x == 1)] = a[(b == 1) & (x == 1)]
    y[(a < 1) & (x == 0)] = np.inf
    y[(b < 1) & (x == 1)] = np.inf
    # Return NaN for out of range parameters
    y[a <= 0] = np.nan
    y[b <= 0] = np.nan
    y[np.isnan(a) | np.isnan(b) | np.isnan(x)] = np.nan
    # Normal values
    k = (a > 0) & (b > 0) & (x > 0) & (x < 1)
    if np.any(k):
        smallx = x[k] < 0.1
        loga = (a[k] - 1) * np.log(x[k])
        logb = np.zeros_like(x[k])
        logb[smallx] = (b[k][smallx] - 1) * np.log1p(-x[k][smallx])
        logb[~smallx] = (b[k][~smallx] - 1) * np.log(1 - x[k][~smallx])
        y[k] = np.exp(loga + logb - betaln(a[k], b[k]))
    return y
