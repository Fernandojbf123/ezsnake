import numpy as np
from typing import Union, Tuple

def lognstat(mu: Union[float, np.ndarray], 
            sigma: Union[float, np.ndarray]) -> Tuple[Union[float, np.ndarray], Union[float, np.ndarray]]:
    """
    Mean and variance for the lognormal distribution.
    
    Parameters:
    -----------
    mu : float or array_like
        Mean of the associated normal distribution
    sigma : float or array_like
        Standard deviation of the associated normal distribution
        
    Returns:
    --------
    m : float or ndarray
        Mean of the lognormal distribution
    v : float or ndarray
        Variance of the lognormal distribution
    """
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    
    sigma = np.where(sigma <= 0, np.nan, sigma)
    
    s2 = sigma ** 2
    
    try:
        m = np.exp(mu + 0.5 * s2)
        v = np.exp(2 * mu + s2) * (np.exp(s2) - 1)
        return m, v
    except Exception:
        raise ValueError("Input size mismatch")
