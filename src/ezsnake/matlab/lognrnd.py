import numpy as np
from typing import Union, Tuple

def lognrnd(mu: Union[float, np.ndarray] = 0, 
           sigma: Union[float, np.ndarray] = 1, 
           size: Union[int, Tuple[int, ...], None] = None) -> np.ndarray:
    """
    Random arrays from the lognormal distribution.
    
    Parameters:
    -----------
    mu : float or array_like, optional
        Mean of the associated normal distribution (default: 0)
    sigma : float or array_like, optional
        Standard deviation of the associated normal distribution (default: 1)
    size : int or tuple of ints, optional
        Output shape. If None, the shape is inferred from mu and sigma
        
    Returns:
    --------
    r : ndarray
        Array of random numbers from the lognormal distribution
    """
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    
    sigma = np.where(sigma < 0, np.nan, sigma)
    
    if size is None:
        size = np.broadcast(mu, sigma).shape
    
    r = np.exp(np.random.randn(*size) * sigma + mu)
    
    return r
