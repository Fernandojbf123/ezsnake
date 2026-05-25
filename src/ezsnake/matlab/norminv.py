import numpy as np
from scipy.stats import norm
from typing import Union

def norminv(p: Union[float, np.ndarray], 
           mu: Union[float, np.ndarray] = 0, 
           sigma: Union[float, np.ndarray] = 1) -> Union[float, np.ndarray]:
    """
    Inverse of the normal cumulative distribution function (cdf).
    
    Parameters:
    -----------
    p : float or array_like
        Probability values
    mu : float or array_like, optional
        Mean (default: 0)
    sigma : float or array_like, optional
        Standard deviation (default: 1)
        
    Returns:
    --------
    x : float or ndarray
        Quantile values
    """
    return norm.ppf(p, loc=mu, scale=sigma)
