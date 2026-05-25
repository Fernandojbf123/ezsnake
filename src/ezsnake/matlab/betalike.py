import numpy as np
from scipy import stats
from scipy.special import psi, gammaln
from scipy.linalg import qr

def betalike(params, data):
    """
    BETALIKE Negative beta log-likelihood function.
    NLOGL = BETALIKE(PARAMS,DATA) returns the negative of beta log-likelihood  
    function for the parameters PARAMS(1) = A and PARAMS(2) = B, given DATA.
    
    [NLOGL, AVAR] = BETALIKE(PARAMS,DATA) returns the inverse of Fisher's
    information matrix, AVAR.  If the input parameter values in PARAMS
    are the maximum likelihood estimates, the diagonal elements of AVAR
    are their asymptotic variances.
    
    The beta distribution is defined on the open interval (0,1).  However, it
    is sometimes also necessary to fit a beta distribution to data that
    include exact zeros or ones.  For such data, the beta likelihood function
    is unbounded, and standard maximum likelihood estimation is not possible.
    In that case, BETALIKE computes a modified likelihood that incorporates the
    zeros or ones by treating them as if they were values that have been
    left-censored at SQRT(REALMIN) or right-censored at 1-EPS/2, respectively.
    
    See also BETAFIT, GAMLIKE, MLE, NORMLIKE, WBLLIKE.
    """
    
    if len(params) != 2:
        raise ValueError('PARAMS must contain exactly 2 parameters')
    
    if not np.isscalar(data) and len(np.asarray(data).shape) > 1:
        if np.asarray(data).shape[0] != 1 and np.asarray(data).shape[1] != 1:
            raise ValueError('DATA must be a vector')
    
    x = np.asarray(data).flatten()
    
    a = params[0]
    b = params[1]
    
    # Return NaN for out of range parameters or data
    if a <= 0:
        a = np.nan
    if b <= 0:
        b = np.nan
    
    xmin = np.min(x)
    xmax = np.max(x)
    x = np.where((0 <= x) & (x <= 1), x, np.nan)
    
    # Separate data into zeros, interior values, and ones
    xl = np.sqrt(np.finfo(float).tiny)  # some tolerance above zero
    xu = 1 - np.finfo(float).eps / 2
    
    if (xl <= xmin) and (xmax <= xu):
        n0 = 0
        n1 = 0
        x2 = x
        n2 = len(x)
    else:
        i0 = (x < xl)
        n0 = np.sum(i0)
        i1 = (x > xu)
        n1 = np.sum(i1)
        i2 = ~(i0 | i1)
        x2 = x[i2]
        n2 = len(x2)
    
    # Compute the usual continuous log-likelihood using values that are
    # strictly within the interval (0,1)
    logx2 = np.log(x2)
    log1mx2 = np.log1p(-x2)
    sumlogx2 = np.sum(logx2)
    sumlog1mx2 = np.sum(log1mx2)
    nlogL = n2 * betaln(a, b) - (a - 1) * sumlogx2 - (b - 1) * sumlog1mx2
    
    # If some values are zero or one, compute a mixed likelihood that includes
    # discrete probabilities for those values.  Note that the asymmetry in xl
    # and xu (relative to 0 and 1, respectively) means that when the vector x
    # contains exact zeros or ones, betalike([a,b],x) is typically not equal
    # to betalike([b,a],1-x).  But that's true even without exact ones and
    # zeros, because of floating point's differing precision at 0 and 1.
    
    # Include F(xl) = Pr(X <= xl) for data that are zeros
    if n0 > 0:
        nlogL = nlogL - n0 * np.log(stats.beta.cdf(xl, a, b))
    
    # Include 1-F(xu) = Pr(X >= xu) for data that are ones
    if n1 > 0:
        nlogL = nlogL - n1 * np.log(1 - stats.beta.cdf(xu, a, b))
    
    # Return only nlogL if avar is not requested
    if True:  # equivalent to nargout > 1 in MATLAB
        if len(data) < 2:
            raise ValueError('Not enough data points to compute variance')
        
        # Compute the Jacobian of the likelihood for values strictly within the
        # interval (0,1)
        J = np.column_stack([logx2 + psi(a + b) - psi(a), 
                            log1mx2 + psi(a + b) - psi(b)])
        
        # Add terms into the Jacobian for the zero and one values
        if n0 > 0 or n1 > 0:
            delta = np.finfo(float).eps ** 0.5
            aa = a + a * delta * np.array([1, -1])
            bb = b + b * delta * np.array([1, -1])
            
            if n0 > 0:
                # Finite central difference approximation to the scores
                # d(F(xl))/d(a,b) for zeros
                da = (np.log(stats.beta.cdf(xl, aa[0], b)) - 
                      np.log(stats.beta.cdf(xl, aa[1], b))) / (2 * a * delta)
                db = (np.log(stats.beta.cdf(xl, a, bb[0])) - 
                      np.log(stats.beta.cdf(xl, a, bb[1]))) / (2 * b * delta)
                J_zeros = np.tile([da, db], (n0, 1))
                J = np.vstack([J, J_zeros])
            
            if n1 > 0:
                # Finite central difference approximation to the scores
                # d(1-F(xu))/d(a,b) for ones
                da = (np.log(1 - stats.beta.cdf(xu, aa[0], b)) - 
                      np.log(1 - stats.beta.cdf(xu, aa[1], b))) / (2 * a * delta)
                db = (np.log(1 - stats.beta.cdf(xu, a, bb[0])) - 
                      np.log(1 - stats.beta.cdf(xu, a, bb[1]))) / (2 * b * delta)
                J_ones = np.tile([da, db], (n1, 1))
                J = np.vstack([J, J_ones])
        
        # Invert the inner product of the Jacobian to get the asymptotic covariance
        Q, R = qr(J, mode='economic')
        if np.any(np.isnan(R)):
            avar = np.array([[np.nan, np.nan], [np.nan, np.nan]])
        else:
            try:
                Rinv = np.linalg.solve(R, np.eye(2))
                avar = Rinv @ Rinv.T
            except np.linalg.LinAlgError:
                avar = np.array([[np.nan, np.nan], [np.nan, np.nan]])
        
        return nlogL, avar
    
    return nlogL


def betaln(a, b):
    """
    Beta function logarithm - equivalent to MATLAB's betaln function
    """
    return gammaln(a) + gammaln(b) - gammaln(a + b)