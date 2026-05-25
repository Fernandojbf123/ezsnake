import numpy as np

def distchck(nparms, *args):
    """
    DISTCHCK Checks the argument list for the probability functions.
    
    Returns:
        errorcode: 0 if sizes are compatible, 1 otherwise
        out_args: tuple of broadcasted arrays (scalars expanded as needed)
    """
    errorcode = 0
    out_args = list(args)
    if nparms == 1:
        return errorcode, tuple(out_args)
    # Check which are scalars
    scalar = [np.isscalar(arg) or (np.size(arg) == 1) for arg in args]
    if all(scalar):
        return errorcode, tuple(out_args)
    # Get shapes
    shapes = [np.shape(arg) for arg in args]
    t = [sh for i, sh in enumerate(shapes) if not scalar[i]]
    size1 = t[0] if t else ()
    # Only check for error
    if len(out_args) == 1:
        for j in range(nparms):
            if not scalar[j] and shapes[j] != size1:
                errorcode = 1
                return errorcode, tuple(out_args)
        return errorcode, tuple(out_args)
    # Expand scalars, check sizes
    for j in range(nparms):
        if scalar[j]:
            vj = args[j]
            arr = np.full(size1, vj)
            out_args[j] = arr
        elif shapes[j] != size1:
            errorcode = 1
            return errorcode, tuple(out_args)
    return errorcode, tuple(out_args)
