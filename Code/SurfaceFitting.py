import xlwings as xw
import utils
import fitting
import ssvi
import phi
from datetime import datetime, date
import numpy as np
import unicodedata as ud

@xw.func
def fit_with_forward_moneyness(dt, dates, money, vol,
                               weight=True,
                               weight_cut = 0.4, calendar_buffer = 2.0e-4,
                               vol_scale =1.0, 
                               max_iter = 10000):
    if type(dates[0]) == str:
        dates = list(map(utils.str2date, dates))

    if type(dt) == datetime:
        dt = dt.date()
        
    if type(dt) == str:
        dt = utils.str2date(dt)

    vol = np.array(vol) * vol_scale
    mult_money = [money for i in range(len(dates))]
    #import pdb;pdb.set_trace()
    fitter = [ssvi.Ssvi([-0.3, 0.01], phi.QuotientPhi([0.4, 0.4])) for i in range(len(dates))]
    surface = fitting.SurfaceFit(dt, dates, mult_money, vol, fitter,
                                 weight = weight,
                                 weight_cut = weight_cut,
                                 calendar_buffer = calendar_buffer)

    surface.calibrate(maxiter = max_iter, verbose = True, method = 'SLSQP')
    surface.visualize()
    
    msg = "¿Quieres ingresar a la base de datos?\n\n"
    msg += "(Si no lo deseas, presiona No,\n"
    msg += "Ajuste los valores de weight_cut | calendar_buffer, etc. e intente ejecutar nuevamente.\n"
    msg += "Los valores predeterminados son weight_cut=0.4 (=40%), calendar_buffer=0.0003.\n"
    msg += "Para reducir la proporción de precios extranjeros, aumente el valor de peso_corte (por ejemplo, 0,7), \n"
    msg += "Cuando ocurre un arbitraje de calendario\n"
    msg += "Establezca calendar_buffer en un valor bajo (no se recomienda menos de 0,00015).\n"
    msg += "[Ejemplo: fit_with_forward_moneyness(dt, dates, money, vol, TRUE, 0.7, 0.0002)] \n"
    msg += "[Si la quinta variable es FALSE, no se da peso]\n"
    msg += "Los valores de las variables 8.ª y 9.ª son el valor de la escala de la vola y \n"
    msg += "(por ejemplo, si los datos vienen como 25.432, configúrelo en 0.01) \n"
    msg += "Este es el valor máximo de iteración durante la optimización\n"
    msg += "(por ejemplo, si el resultado no es satisfactorio, configúrelo en 20000).\n\n"
    msg += "¡Precaución! Si la opción de cálculo de Excel es automática, se ejecutará automáticamente cuando cambie el valor de los datos.\n"
    
    m_res = utils.Mbox("", msg, 4)
    if m_res == 6:
        rho = ud.lookup("GREEK SMALL LETTER RHO")
        theta = ud.lookup("GREEK SMALL LETTER THETA")
        eta = ud.lookup("GREEK SMALL LETTER ETA")
        gamma = ud.lookup("GREEK SMALL LETTER GAMMA")
        res = [[None, rho, theta, eta, gamma]]
        _params =surface.params
        
        for i, p in enumerate(_params):
            row = [dates[i]] + list(p)
            res = res + [row]
        return res
    else:
        return

    
    

#import xlwings as xw;xw.serve()
    
