import numpy as np
from typing import Union, Tuple, Optional
from .normlike import normlike

def lognlike(params: Union[list, np.ndarray], 
            data: np.ndarray,
            censoring: Optional[np.ndarray] = None,
            freq: Optional[np.ndarray] = None) -> Union[float, Tuple[float, np.ndarray]]:
    """
    Negative log-likelihood for the lognormal distribution.
    
    Parameters:
    -----------
    params : array_like
        Parameters [mu, sigma] where mu and sigma are parameters of the 
        associated normal distribution
    data : array_like
        Data values (must be positive)
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
    
    if np.any(data < 0):
        return np.inf
    
    if censoring is not None or freq is not None:
        raise NotImplementedError("Censoring and frequency weights not implemented")
    
    logdata = np.log(data)
    
    nlogL = normlike(params, logdata, censoring, freq)
    
    if freq is None:
        freq = np.ones_like(data)
    if censoring is None:
        censoring = np.zeros_like(data, dtype=bool)
    
    nlogL = nlogL + np.sum(freq * logdata * (1 - censoring))
    
    return nlogL
