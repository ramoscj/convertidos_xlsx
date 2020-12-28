from sys import platform
import platform

class PersonaFuga:

    def __init__(self, crr: int, fuga: int, stock: int, rut: str):
        self.crr = crr
        self.fuga = fuga
        self.stock = stock
        self.rut = rut
        if type(self.crr) is not int:
            raise Exception('El valor "crr" no es un int')
        elif type(self.fuga) is not int:
            raise Exception('El valor "fuga" no es un int')
        elif type(self.stock) is not int:
            raise Exception('El valor "stock" no es un int')
        elif type(self.rut) is not str:
            raise Exception('El valor "rut" no es un str')

# jorge = PersonaFuga(1,2,1,2)

from itertools import cycle

def digito_verificador(rut):
    reversed_digits = map(int, reversed(str(rut)))
    factors = cycle(range(2, 8))
    s = sum(d * f for d, f in zip(reversed_digits, factors))
    return (-s) % 11

print(digito_verificador(264295645))