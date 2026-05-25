import numpy as np
from typing import Union, Optional

def lognpdf(x: Union[float, np.ndarray], 
           mu: Union[float, np.ndarray] = 0, 
           sigma: Union[float, np.ndarray] = 1) -> Union[float, np.ndarray]:
    """
    Lognormal probability density function (pdf).
    
    Parameters:
    -----------
    x : float or array_like
        Values at which to evaluate the pdf
    mu : float or array_like, optional
        Mean of the associated normal distribution (default: 0)
    sigma : float or array_like, optional  
        Standard deviation of the associated normal distribution (default: 1)
        
    Returns:
    --------
    y : float or ndarray
        Probability density values
    """
    x = np.asarray(x)
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    
    sigma = np.where(sigma <= 0, np.nan, sigma)
    
    x = np.where(x <= 0, np.inf, x)
    
    try:
        y = np.exp(-0.5 * ((np.log(x) - mu) / sigma) ** 2) / (x * np.sqrt(2 * np.pi) * sigma)
        return y
    except Exception:
        raise ValueError("Input size mismatch")
