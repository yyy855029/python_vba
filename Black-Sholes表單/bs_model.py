import xlwings as xw
import xloil
import numpy as np
from scipy.stats import norm
import matplotlib.pyplot as plt
import matplotlib.style as style

style.use('seaborn')

#設定字體和解決圖像中文顯示問題
plt.rcParams['font.sans-serif']=['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus']=False


@xloil.func
def bs(callput_flag,S,K,T,r,sigma):
    d1=(np.log(S/K)+(r+0.5*sigma**2)*T)/(sigma*np.sqrt(T))
    d2=d1-sigma*np.sqrt(T)
   
    if callput_flag=='Call':
        symbol=1
    elif callput_flag=='Put':
        symbol=-1
    
    price=symbol*S*norm.cdf(symbol*d1,0,1)-symbol*K*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)
    delta=symbol*norm.cdf(symbol*d1,0,1)
    gamma=norm.pdf(d1,0,1)/(S*sigma*np.sqrt(T))
    theta=-S*norm.pdf(symbol*d1,0,1)*sigma/(2*np.sqrt(T))-symbol*r*K*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)
    vega=K*np.exp(-r*T)*norm.pdf(d2,0,1)*np.sqrt(T)
    rho=symbol*K*T*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)
    
    return price,delta,gamma,theta,vega,rho


@xloil.func
def call_bs():
    wb=xw.Book.caller()
    sht=wb.sheets[0]
    
    callput_flag=sht.range('C3').value
    r=sht.range('C7').value
    S=np.array(sht.range('B17:B37').value)
    K=np.array(sht.range('C17:C37').value)
    T=np.array(sht.range('D17:D37').value)
    sigma=np.array(sht.range('F17:F37').value)
    price,delta,gamma,theta,vega,rho=bs(callput_flag,S,K,T,r,sigma)
    sht.range('G17').options(transpose=True).value=price
    sht.range('I17').options(transpose=True).value=delta
    sht.range('J17').options(transpose=True).value=gamma
    sht.range('K17').options(transpose=True).value=theta
    sht.range('L17').options(transpose=True).value=vega
    sht.range('M17').options(transpose=True).value=rho

    X_values=sht.range('P17:P37').value
    Y_values=sht.range('Q17:Q37').value
    X_label=sht.range('P16').value
    Y_label=sht.range('Q16').value

    fig,ax=plt.subplots(1,figsize=(8,3))
    ax.plot(X_values,Y_values)
    ax.set_title('Black-Scholes Model and Greeks',fontsize=14)
    ax.set_xlabel(X_label,fontsize=10)
    ax.set_ylabel(Y_label,fontsize=10)

    sht.pictures.add(fig,name='Plot',update=True,
        left=sht.range('F2').left,top=sht.range('F2').top)







