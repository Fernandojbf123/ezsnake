import numpy as np
from scipy import stats, optimize
from scipy.special import gammaln, psi, logsumexp
from .betalike import betalike
from .norminv import norminv

def betafit(x, alpha=0.05):
    """
    BETAFIT Parameter estimates and confidence intervals for beta distributed data.
    BETAFIT(X) Returns the maximum likelihood estimates of the parameters
    of the beta distribution given the data in the vector, X.
    
    [PHAT, PCI] = BETAFIT(X,ALPHA) gives MLEs and 100(1-ALPHA) percent
    confidence intervals given the data. By default, the optional parameter
    ALPHA = 0.05 corresponding to 95% confidence intervals.
    
    The beta distribution is defined on the open interval (0,1).  However, it
    is sometimes also necessary to fit a beta distribution to data that
    include exact zeros or ones.  For such data, the beta likelihood function
    is unbounded, and standard maximum likelihood estimation is not possible.
    In that case, BETAFIT maximizes a modified likelihood that incorporates
    the zeros or ones by treating them as if they were values that have been
    left-censored at SQRT(REALMIN) or right-censored at 1-EPS/2, respectively.
    
    See also BETACDF, BETAINV, BETALIKE, BETAPDF, BETARND, BETASTAT, MLE.
    
    Reference:
       (1994) Hahn, Gerald J., and Shapiro, Samuel, S. "Statistical Models in
       Engineering", Wiley Classics Library, John Wiley & Sons, p. 95.
    """
    
    alpha = float(alpha)
    
    if len(x) == 0:
        phat = np.array([np.nan, np.nan])
        pci = np.array([[np.nan, np.nan], [np.nan, np.nan]])
        return phat, pci
    
    x = np.asarray(x).flatten()
    
    # Remove missing values from the data
    x = x[~np.isnan(x)]
    
    # Cannot fit data outside of the closed interval [0,1], or constant data
    xmin = np.min(x)
    xmax = np.max(x)
    
    if (xmin < 0) or (xmax > 1) or not np.all(np.isreal(x)):
        raise ValueError('X must be real values in the interval [0,1]')
    
    if abs(xmin - xmax) <= 2 * np.finfo(float).eps * xmax:
        raise ValueError('X must contain distinct values')
    
    # Initial parameter estimates
    n = len(x)
    sumlogx = np.sum(np.log(x))
    sumlog1mx = np.sum(np.log1p(-x))
    tmp1 = np.exp(sumlog1mx / n)
    tmp2 = np.exp(sumlogx / n)
    
    tmp3 = (1 - tmp1 - tmp2)
    ahat = 0.5 * (1 - tmp1) / tmp3
    bhat = 0.5 * (1 - tmp2) / tmp3
    pstart = np.array([ahat, bhat])
    
    # If all values are strictly within the interval (0,1), use
    # maximum likelihood with the usual continuous log-likelihood
    xl = np.sqrt(np.finfo(float).tiny)  # some tolerance above zero
    xu = 1 - np.finfo(float).eps / 2
    
    if (xl <= xmin) and (xmax <= xu):
        def negloglike_cts(p):
            p = np.exp(p)  # remove log transform
            return n * betaln(p[0], p[1]) - (p[0] - 1) * sumlogx - (p[1] - 1) * sumlog1mx
        
        negloglike = negloglike_cts
    else:
        # If some values are zero or one, maximize a mixed likelihood that
        # includes discrete probabilities for those values
        x0 = (x < xl)
        n0 = np.sum(x0)
        x1 = (x > xu)
        n1 = np.sum(x1)
        x2 = x[~x0 & ~x1]
        n2 = len(x2)
        sumlogx2 = np.sum(np.log(x2))
        sumlog1mx2 = np.sum(np.log1p(-x2))
        
        def negloglike_mixed(p):
            p = np.exp(p)  # remove log transform
            nll = n2 * betaln(p[0], p[1]) - (p[0] - 1) * sumlogx2 - (p[1] - 1) * sumlog1mx2
            
            # Include F(xl) = Pr(X <= xl) for data that are zeros
            if n0 > 0:
                nll = nll - n0 * np.log(stats.beta.cdf(xl, p[0], p[1]))
            
            # Include 1-F(xu) = Pr(X >= xu) for data that are ones
            if n1 > 0:
                nll = nll - n1 * np.log(1 - stats.beta.cdf(xu, p[0], p[1]))
            
            return nll
        
        negloglike = negloglike_mixed
    
    # Maximize the likelihood using a log transform for the parameters, to ensure
    # the parameters are positive
    pstart = np.log(pstart)
    result = optimize.minimize(negloglike, pstart, method='Nelder-Mead',
                             options={'xatol': 1e-7, 'fatol': 1e-7})
    phat = np.exp(result.x)
    
    if alpha is not None:
        
        # Compute CIs on the log scale for both params
        _, acov = betalike(phat, x)
        logphat = np.log(phat)
        selog = np.sqrt(np.diag(acov)) / phat
        
        p_int = np.array([alpha / 2, 1 - alpha / 2])
        pci = np.exp(norminv(np.column_stack([p_int, p_int]), 
                            np.row_stack([logphat, logphat]), 
                            np.row_stack([selog, selog])))
        return phat, pci
    else:
        return phat


def betaln(a, b):
    """
    Beta function logarithm - equivalent to MATLAB's betaln function
    """
    return gammaln(a) + gammaln(b) - gammaln(a + b)