class SsviPhi:        
    def __call__(self, j):
        pass

class QuotientPhi(SsviPhi):
    def __init__(self, x):
        self.eta  = x[0]
        self.gamma= x[1]

    def __call__(self, theta):
        ret = self.eta
        ret /= theta**self.gamma
        ret /= (1.+theta)**(1.0-self.gamma)
        return ret
