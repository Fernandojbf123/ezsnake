import numpy as np
from scipy.special import erfc
from typing import Union

def logncdf(x: Union[float, np.ndarray], 
           mu: Union[float, np.ndarray] = 0, 
           sigma: Union[float, np.ndarray] = 1,
           upper: bool = False) -> Union[float, np.ndarray]:
    """
    Lognormal cumulative distribution function (cdf).
    
    Parameters:
    -----------
    x : float or array_like
        Values at which to evaluate the cdf
    mu : float or array_like, optional
        Mean of the associated normal distribution (default: 0)
    sigma : float or array_like, optional
        Standard deviation of the associated normal distribution (default: 1)
    upper : bool, optional
        If True, returns upper tail probability (default: False)
        
    Returns:
    --------
    p : float or ndarray
        Cumulative probability values
    """
    x = np.asarray(x)
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    
    sigma = np.where(sigma <= 0, np.nan, sigma)
    
    x = np.where(x < 0, 0, x)
    
    try:
        z = (np.log(x) - mu) / sigma
        if upper:
            z = -z
    except Exception:
        raise ValueError("Input size mismatch")
    
    p = 0.5 * erfc(-z / np.sqrt(2))
    
    return p
