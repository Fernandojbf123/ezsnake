import numpy as np
from scipy.stats import norm, chi2
from typing import Union, Tuple, Optional

def normfit(x: np.ndarray, 
           alpha: float = 0.05,
           censoring: Optional[np.ndarray] = None,
           freq: Optional[np.ndarray] = None) -> Tuple[float, float, Optional[Tuple[float, float]], Optional[Tuple[float, float]]]:
    """
    Parameter estimates and confidence intervals for normal data.
    
    Parameters:
    -----------
    x : array_like
        Data values
    alpha : float, optional
        Significance level for confidence intervals (default: 0.05)
    censoring : array_like, optional
        Boolean array indicating censored observations
    freq : array_like, optional
        Frequency weights
        
    Returns:
    --------
    muhat : float
        Estimated mean
    sigmahat : float
        Estimated standard deviation
    muci : tuple, optional
        Confidence interval for mean (if requested)
    sigmaci : tuple, optional
        Confidence interval for standard deviation (if requested)
    """
    x = np.asarray(x)
    
    if np.any(~np.isfinite(x)):
        raise ValueError("Data contains non-finite values")
    
    if censoring is not None or freq is not None:
        raise NotImplementedError("Censoring and frequency weights not implemented")
    
    muhat = np.mean(x)
    sigmahat = np.std(x, ddof=1)
    
    n = len(x)
    se_mu = sigmahat / np.sqrt(n)
    
    t_crit = norm.ppf(1 - alpha/2)
    muci = (muhat - t_crit * se_mu, muhat + t_crit * se_mu)
    
    chi2_low = n - 1
    chi2_high = n - 1
    
    sigmaci_low = np.sqrt((n-1) * sigmahat**2 / chi2.ppf(1 - alpha/2, n-1))
    sigmaci_high = np.sqrt((n-1) * sigmahat**2 / chi2.ppf(alpha/2, n-1))
    sigmaci = (sigmaci_low, sigmaci_high)
    
    return muhat, sigmahat, muci, sigmaci
