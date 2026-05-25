import numpy as np
from scipy.special import erfc
from .distchck import distchck
from .norminv import norminv

def normcdf(x, mu=0, sigma=1, pcov=None, alpha=0.05, tail='lower'):
    """
    NORMCDF Normal cumulative distribution function (cdf).
    P = NORMCDF(X,MU,SIGMA) returns the cdf of the normal distribution with
    mean MU and standard deviation SIGMA, evaluada en los valores de X.
    El tamaño de P es el tamaño común de X, MU y SIGMA. Un escalar de entrada
    funciona como una matriz constante del mismo tamaño que las otras entradas.

    Valores por defecto para MU y SIGMA son 0 y 1, respectivamente.

    [P,PLO,PUP] = NORMCDF(X,MU,SIGMA,PCOV,ALPHA) produce límites de confianza
    para P cuando los parámetros MU y SIGMA son estimaciones. PCOV es una
    matriz 2x2 con la covarianza de los parámetros estimados. ALPHA es el
    nivel de significancia (por defecto 0.05).

    tail='upper' calcula la probabilidad de cola superior.
    """
    # Manejo de argumentos y tamaños
    x = np.asarray(x)
    mu = np.asarray(mu)
    sigma = np.asarray(sigma)
    errorcode, (x, mu, sigma) = distchck(3, x, mu, sigma)
    if errorcode > 0:
        raise ValueError('Input size mismatch')
    z = (x - mu) / sigma
    if tail == 'upper':
        z = -z
    p = np.full_like(z, np.nan, dtype=float)
    # Casos sigma=0
    if tail == 'upper':
        p[(sigma == 0) & (x < mu)] = 1
        p[(sigma == 0) & (x >= mu)] = 0
    else:
        p[(sigma == 0) & (x < mu)] = 0
        p[(sigma == 0) & (x >= mu)] = 1
    # Casos normales
    mask = sigma > 0
    p[mask] = 0.5 * erfc(-z[mask] / np.sqrt(2))
    # Límites de confianza si se solicita
    plo = pup = None
    if pcov is not None:
        if not (isinstance(pcov, np.ndarray) and pcov.shape == (2,2)):
            raise ValueError('pcov debe ser una matriz 2x2')
        if not (0 < alpha < 1):
            raise ValueError('alpha debe estar entre 0 y 1')
        z = z[mask].reshape(-1)
        sigma_mask = sigma[mask].reshape(-1)
        zvar = (pcov[0,0] + 2*pcov[0,1]*z + pcov[1,1]*z**2) / (sigma_mask**2)
        if np.any(zvar < 0):
            raise ValueError('Covarianza no es semidefinida positiva')
        normz = -norminv(alpha/2)
        halfwidth = normz * np.sqrt(zvar)
        zlo = z - halfwidth
        zup = z + halfwidth
        plo = np.full_like(p, np.nan, dtype=float)
        pup = np.full_like(p, np.nan, dtype=float)
        plo[mask] = 0.5 * erfc(-zlo / np.sqrt(2))
        pup[mask] = 0.5 * erfc(-zup / np.sqrt(2))
    if plo is not None and pup is not None:
        return p, plo, pup
    return p
