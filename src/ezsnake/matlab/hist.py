import numpy as np

def hist(y, x=None):
    """
    Histogram bin counts similar to MATLAB's hist for numeric data.
    If x is None, uses 10 bins.
    If x is an integer, uses x bins.
    If x is a vector, uses it as bin centers (MATLAB style).
    Returns (counts, bin_centers)
    """
    y = np.asarray(y).flatten()
    if x is None:
        bins = 10
        counts, bin_edges = np.histogram(y, bins=bins)
        bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2
    elif np.isscalar(x):
        bins = int(x)
        counts, bin_edges = np.histogram(y, bins=bins)
        bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2
    else:
        # x is a vector of bin centers (MATLAB style)
        x = np.asarray(x)
        # Compute bin edges from centers
        bin_edges = np.zeros(len(x) + 1)
        bin_edges[1:-1] = (x[:-1] + x[1:]) / 2
        bin_edges[0] = -np.inf
        bin_edges[-1] = np.inf
        counts, _ = np.histogram(y, bins=bin_edges)
        bin_centers = x
    return counts, bin_centers
