import sys, json; sys.path.insert(0, '.')
import numpy as np, packdata as PD
def conv(o):
    if isinstance(o, (np.floating, np.integer)): return o.item()
    if isinstance(o, np.ndarray): return o.tolist()
    return str(o)
def export(k, T, path):
    D = PD.build(k, T)
    json.dump(D, open(path, 'w'), default=conv)
    return D
if __name__ == '__main__':
    export(sys.argv[1], int(sys.argv[2]), sys.argv[3]); print('ok')
