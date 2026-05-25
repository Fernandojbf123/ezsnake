import numpy as np
from typing import Tuple, Optional
from .normfit import normfit

def lognfit(x: np.ndarray, 
           alpha: float = 0.05,
           censoring: Optional[np.ndarray] = None,
           freq: Optional[np.ndarray] = None) -> Tuple[np.ndarray, Optional[np.ndarray]]:
    """
    Parameter estimates and confidence intervals for lognormal data.
    
    Parameters:
    -----------
    x : array_like
        Data values (must be positive)
    alpha : float, optional
        Significance level for confidence intervals (default: 0.05)
    censoring : array_like, optional
        Boolean array indicating censored observations
    freq : array_like, optional
        Frequency weights
        
    Returns:
    --------
    parmhat : ndarray
        Parameter estimates [mu, sigma]
    parmci : ndarray, optional
        Confidence intervals for parameters
    """
    x = np.asarray(x)
    
    if not np.all(x > 0):
        raise ValueError("All data values must be positive")
    
    if censoring is not None or freq is not None:
        raise NotImplementedError("Censoring and frequency weights not implemented")
    
    log_x = np.log(x)
    
    muhat, sigmahat, muci, sigmaci = normfit(log_x, alpha)
    
    parmhat = np.array([muhat, sigmahat])
    
    if muci is not None and sigmaci is not None:
        parmci = np.array([list(muci), list(sigmaci)])
        return parmhat, parmci
    else:
        return parmhat, None
