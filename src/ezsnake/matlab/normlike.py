import numpy as np
from typing import Union, Tuple, Optional

def normlike(params: Union[list, np.ndarray], 
            data: np.ndarray,
            censoring: Optional[np.ndarray] = None,
            freq: Optional[np.ndarray] = None) -> Union[float, Tuple[float, np.ndarray]]:
    """
    Negative log-likelihood for the normal distribution.
    
    Parameters:
    -----------
    params : array_like
        Parameters [mu, sigma]
    data : array_like
        Data values
    censoring : array_like, optional
        Boolean array indicating censored observations
    freq : array_like, optional
        Frequency weights
        
    Returns:
    --------
    nlogL : float
        Negative log-likelihood
    avar : ndarray, optional
        Asymptotic variance matrix (if requested)
    """
    params = np.asarray(params)
    data = np.asarray(data)
    
    if len(params) != 2:
        raise ValueError("Parameters must have length 2")
    
    mu, sigma = params
    
    if sigma <= 0:
        return np.inf
    
    if censoring is not None or freq is not None:
        raise NotImplementedError("Censoring and frequency weights not implemented")
    
    n = len(data)
    
    log_likelihood = -0.5 * n * np.log(2 * np.pi) - n * np.log(sigma) - 0.5 * np.sum((data - mu)**2) / sigma**2
    
    nlogL = -log_likelihood
    
    return nlogL
