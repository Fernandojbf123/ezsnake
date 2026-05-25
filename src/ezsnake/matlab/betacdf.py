import numpy as np
from scipy.stats import beta

def betacdf(x, a, b, tail='lower'):
    """
    BETACDF Beta cumulative distribution function.
    P = BETACDF(X,A,B) returns the beta cumulative distribution
    function with parameters A and B at the values in X.

    The size of P is the common size of the input arguments. A scalar input  
    functions as a constant matrix of the same size as the other inputs.    

    BETAINC does the computational work.

    P = BETACDF(X,A,B,'upper') returns the upper tail probability of the beta 
    distribution function with parameters A and B at the values in X.

    See also BETAFIT, BETAINV, BETALIKE, BETAPDF, BETARND, BETASTAT, CDF,
             BETAINC.

    Reference:
       [1]  M. Abramowitz and I. A. Stegun, "Handbook of Mathematical
       Functions", Government Printing Office, 1964, 26.5.
    """
    x = np.asarray(x)
    a = np.asarray(a)
    b = np.asarray(b)
    # Broadcast to common shape
    x, a, b = np.broadcast_arrays(x, a, b)
    p = np.full_like(x, np.nan, dtype=float)
    okAB = (0 < a) & (a < np.inf) & (0 < b) & (b < np.inf)
    k = okAB & (0 <= x) & (x <= 1)
    if tail == 'upper':
        p[okAB & (x <= 0)] = 1.0
        p[okAB & (x >= 1)] = 0.0
        if np.any(k):
            p[k] = beta.sf(x[k], a[k], b[k])
    else:
        p[okAB & (x < 0)] = 0.0
        p[okAB & (x > 1)] = 1.0
        if np.any(k):
            p[k] = beta.cdf(x[k], a[k], b[k])
    return p
