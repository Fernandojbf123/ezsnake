import numpy as np

def normpdf(x, mu=0, sigma=1):
    """
    NORMPDF Normal probability density function (pdf).
    Y = NORMPDF(X,MU,SIGMA) returns the pdf of the normal distribution with
    mean MU and standard deviation SIGMA, evaluated at the values in X.
    The size of Y is the common size of the input arguments.  A scalar
    input functions as a constant matrix of the same size as the other
    inputs.

    Default values for MU and SIGMA are 0 and 1 respectively.

    See also NORMCDF, NORMFIT, NORMINV, NORMLIKE, NORMRND, NORMSTAT.

    References:
       [1] Evans, M., Hastings, N., and Peacock, B. (1993) Statistical
           Distributions, 2nd ed., Wiley, 170pp.
    """
    x = np.asarray(x)
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    # Return NaN for out of range parameters.
    sigma = np.where(sigma <= 0, np.nan, sigma)
    try:
        y = np.exp(-0.5 * ((x - mu) / sigma) ** 2) / (np.sqrt(2 * np.pi) * sigma)
    except Exception:
        raise ValueError('Input size mismatch')
    return y
