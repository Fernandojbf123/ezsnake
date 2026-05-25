import numpy as np
import warnings
from scipy.special import betaincinv, betainc
from .distchck import distchck
from .betacdf import betacdf

def betainv(p, a, b):
    """
    BETAINV Inverse of the beta cumulative distribution function (cdf).
    X = BETAINV(P,A,B) returns the inverse of the beta cdf with 
    parameters A and B at the values in P.

    The size of X is the common size of the input arguments. A scalar input  
    functions as a constant matrix of the same size as the other inputs.    

    BETAINV uses Newton's method to converge to the solution (en MATLAB).
    En Python, se usa scipy.special.betaincinv.

    See also BETACDF, BETAFIT, BETALIKE, BETAPDF, BETARND, BETASTAT, ICDF.
    """
    if a is None or b is None:
        raise ValueError('Too few inputs')
    errorcode, (p, a, b) = distchck(3, p, a, b)
    if errorcode > 0:
        raise ValueError('Input size mismatch')
    p = np.asarray(p)
    a = np.asarray(a)
    b = np.asarray(b)
    # Weed out any out of range parameters or probabilities.
    okAB = (0 < a) & (a < np.inf) & (0 < b) & (b < np.inf)
    k = okAB & (0 <= p) & (p <= 1)
    allOK = np.all(k)
    # Fill in NaNs for out of range cases.
    x = np.full(k.shape, np.nan, dtype=float)
    if not allOK:
        if np.any(k):
            if p.size > 1:
                p2 = p[k]
            else:
                p2 = p
            if a.size > 1:
                a2 = a[k]
            else:
                a2 = a
            if b.size > 1:
                b2 = b[k]
            else:
                b2 = b
        else:
            return x
    else:
        p2, a2, b2 = p, a, b
    # Call betaincinv to find a root of betainc(q,a,b) = p
    q = betaincinv(a2, b2, p2)
    delta = np.finfo(float).eps
    badcdf = (np.abs(betainc(a2, b2, q) - p2) / np.maximum(p2, 1e-12)) > np.sqrt(delta)
    if np.any(badcdf):
        # Intentar ajustar con betacdf
        idx = np.where(badcdf)[0]
        for i in idx:
            q_i = q[i]
            a_i = a2[i] if a2.size > 1 else a2
            b_i = b2[i] if b2.size > 1 else b2
            p_i = p2[i] if p2.size > 1 else p2
            q_minus = q_i - delta
            q_plus = q_i + delta
            cdf_minus = betacdf(q_minus, a_i, b_i)
            cdf_plus = betacdf(q_plus, a_i, b_i)
            if (cdf_minus - p_i) * (cdf_plus - p_i) > 0:
                warnings.warn(f'betainv: NoConvergence for a={a_i}, b={b_i}, p={p_i}')
    # Broadcast the values to the correct place if need be.
    if allOK:
        x = q
    else:
        x[k] = q
    return x
