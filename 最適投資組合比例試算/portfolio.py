import xlwings as xw
import xloil
import numpy as np
from scipy.optimize import minimize


@xloil.func
def portfolio_mean(weight_array,mean_array):
	return np.sum(mean_array*weight_array)


@xloil.func
def portfolio_std(weight_array,cov_array):
	return np.sqrt(np.dot(weight_array.T,np.dot(cov_array,weight_array)))


@xloil.func
def ef_max_mean(weight_array,mean_array,cov_array,tolerance_risk):
	init_guess=np.ones(len(weight_array))*(1.0/len(weight_array))
	bnds=tuple((0,1) for x in range(len(weight_array)))
	cons=({'type':'eq','fun':lambda x:np.sum(x)-1},
		  {'type':'ineq','fun':lambda x:tolerance_risk-portfolio_std(x,cov_array)})

	fun=lambda x,y:-portfolio_mean(x,y)

	result=minimize(fun,init_guess,args=(mean_array),method='SLSQP',bounds=bnds,constraints=cons)

	return result['x'].round(5)


@xloil.func
def ef_min_var(weight_array,mean_array,cov_array,tolerance_return):
    init_guess=np.ones(len(weight_array))*(1.0/len(weight_array))
    bnds=tuple((0,1) for x in range(len(weight_array)))
    cons=({'type':'eq','fun':lambda x:np.sum(x)-1},
		  {'type':'ineq','fun':lambda x:portfolio_mean(x,mean_array)-tolerance_return})
    
    result=minimize(portfolio_std,init_guess,args=(cov_array),method='SLSQP',bounds=bnds,constraints=cons)
    
    return result['x'].round(5)


@xloil.func    
def ef_sharpe(weight_array,mean_array,cov_array,rf):
    init_guess=np.ones(len(weight_array))*(1.0/len(weight_array))
    bnds=tuple((0,1) for x in range(len(weight_array)))
    cons=({'type':'eq','fun':lambda x:np.sum(x)-1},
          {'type':'ineq','fun':lambda x:portfolio_std(x,cov_array)},
          {'type':'ineq','fun':lambda x:portfolio_mean(x,mean_array)-rf})

    fun=lambda x,y,z:-(portfolio_mean(x,y)-rf)/portfolio_std(x,z)
    
    result=minimize(fun,init_guess,args=(mean_array,cov_array),method='SLSQP',bounds=bnds,constraints=cons)

    return result['x'].round(5)


@xloil.func
def call_max_mean():
    wb=xw.Book.caller()
    sht=wb.sheets[0]

    mean_array=np.array(sht.range('J10:J12').value)
    weight_array=np.array(sht.range('E10:E12').value)
    cov_array=np.array(sht.range('M10:O12').value)
    tolerance_risk=sht.range('E3').value
    ef_max_mean_weights=ef_max_mean(weight_array,mean_array,cov_array,tolerance_risk)
    sht.range('E10').options(transpose=True).value=ef_max_mean_weights
    weight_array=np.array(sht.range('E10:E12').value)
    sht.range('C3').value=portfolio_mean(weight_array,mean_array)
    sht.range('C4').value=portfolio_std(weight_array,cov_array)


@xloil.func
def call_min_var():
    wb=xw.Book.caller()
    sht=wb.sheets[0]

    mean_array=np.array(sht.range('J10:J12').value)
    weight_array=np.array(sht.range('E10:E12').value)
    cov_array=np.array(sht.range('M10:O12').value)
    tolerance_return=sht.range('E4').value
    ef_min_var_weights=ef_min_var(weight_array,mean_array,cov_array,tolerance_return)
    sht.range('E10').options(transpose=True).value=ef_min_var_weights
    weight_array=np.array(sht.range('E10:E12').value)
    sht.range('C3').value=portfolio_mean(weight_array,mean_array)
    sht.range('C4').value=portfolio_std(weight_array,cov_array)


@xloil.func
def call_sharpe():
    wb=xw.Book.caller()
    sht=wb.sheets[0]

    mean_array=np.array(sht.range('J10:J12').value)
    weight_array=np.array(sht.range('E10:E12').value)
    cov_array=np.array(sht.range('M10:O12').value)
    rf=sht.range('J7').value
    ef_sharpe_weights=ef_sharpe(weight_array,mean_array,cov_array,rf)
    sht.range('E10').options(transpose=True).value=ef_sharpe_weights
    weight_array=np.array(sht.range('E10:E12').value)
    sht.range('C3').value=portfolio_mean(weight_array,mean_array)
    sht.range('C4').value=portfolio_std(weight_array,cov_array)






