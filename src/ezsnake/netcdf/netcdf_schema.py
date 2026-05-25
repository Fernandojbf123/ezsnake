
from typing import Any
import numpy as np

class BuildNetCDFVariable:
    def __init__(self, value: np.ndarray = np.array([]), dims: list[str]= [], attrs: dict[str, Any]={}):
        self.value = value or np.array([])
        self.dims = dims or []
        self.attrs = attrs or {}
        
    def set_value(self, value):
        self.value = value
        
    def set_dims(self, dims):
        """ dims = ['lon', 'lat'] """
        self.dims = dims
        
    def set_attrs(self, attrs):
        """ attrs = {'_FillValue': -9999.0, 'units': 'K', 'long_name': 'Temperatura'} """
        self.attrs.update(attrs)
        
    def add_attr(self, attr_name, attr_value):
        self.attrs[attr_name] = attr_value
        
    def get_info(self):
        return {
            'value': self.value,
            'dims': self.dims,
            'attrs': self.attrs
        }