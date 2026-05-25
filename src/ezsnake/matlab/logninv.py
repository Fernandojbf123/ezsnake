import numpy as np
from scipy.special import erfcinv
from typing import Union

def logninv(p: Union[float, np.ndarray], 
           mu: Union[float, np.ndarray] = 0, 
           sigma: Union[float, np.ndarray] = 1) -> Union[float, np.ndarray]:
    """
    Inverse of the lognormal cumulative distribution function (cdf).
    
    Parameters:
    -----------
    p : float or array_like
        Probability values
    mu : float or array_like, optional
        Mean of the associated normal distribution (default: 0)
    sigma : float or array_like, optional
        Standard deviation of the associated normal distribution (default: 1)
        
    Returns:
    --------
    x : float or ndarray
        Quantile values
    """
    p = np.asarray(p)
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    
    sigma = np.where(sigma <= 0, np.nan, sigma)
    p = np.where((p < 0) | (p > 1), np.nan, p)
    
    logx0 = -np.sqrt(2) * erfcinv(2 * p)
    
    try:
        x = np.exp(sigma * logx0 + mu)
        return x
    except Exception:
        raise ValueError("Input size mismatch")
