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
def gbs(callput_flag,S,K,T,r,b,sigma):
    d1=(np.log(S/K)+(b+0.5*sigma**2)*T)/(sigma*np.sqrt(T))
    d2=d1-sigma*np.sqrt(T)
   
    if callput_flag=='Call':
        symbol=1
    elif callput_flag=='Put':
        symbol=-1
    
    price=symbol*S*np.exp((b-r)*T)*norm.cdf(symbol*d1,0,1)-symbol*K*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)
    delta=symbol*np.exp((b-r)*T)*norm.cdf(symbol*d1,0,1)
    gamma=np.exp((b-r)*T)*norm.pdf(d1,0,1)/(S*sigma*np.sqrt(T))
    theta=-S*np.exp((b-r)*T)*norm.pdf(symbol*d1,0,1)*sigma/(2*np.sqrt(T))-symbol*r*K*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)-symbol*(b-r)*S*np.exp((b-r)*T)*norm.cdf(symbol*d1,0,1)
    vega=K*np.exp(-r*T)*norm.pdf(d2,0,1)*np.sqrt(T)
    
    if b==0:
        rho=-T*price
    else:
        rho=symbol*K*T*np.exp(-r*T)*norm.cdf(symbol*d2,0,1)
    
    return price,delta,gamma,theta,vega,rho


@xloil.func
def bs_newton_iv(callput_flag,market_price,S,K,T,r,sigma):
    error=0.00001
    error_count=0

    sigma_temp=(abs(np.log(S/K)+r*T)*2/T)**0.5
    f,_,_,_,vega_temp,_=bs(callput_flag,S,K,T,r,sigma_temp)

    while abs(market_price-f)>error and error_count<100:
        sigma_temp+=(market_price-f)/vega_temp
        f,_,_,_,vega_temp,_=bs(callput_flag,S,K,T,r,sigma_temp)
        
        error_count+=1

    return sigma_temp


@xloil.func
def gbs_newton_iv(callput_flag,market_price,S,K,T,r,b,sigma):
    error=0.00001
    error_count=0

    sigma_temp=(abs(np.log(S/K)+r*T)*2/T)**0.5
    f,_,_,_,vega_temp,_=gbs(callput_flag,S,K,T,r,b,sigma_temp)

    while abs(market_price-f)>error and error_count<100:
        sigma_temp+=(market_price-f)/vega_temp
        f,_,_,_,vega_temp,_=gbs(callput_flag,S,K,T,r,b,sigma_temp)
        
        error_count+=1

    return sigma_temp


@xloil.func
def bs_bisection_iv(callput_flag,market_price,S,K,T,r,sigma):
    error=0.00001
    error_count=0

    sigma_L=0.001
    sigma_H=10
    
    fL,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_L)
    fH,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_H)
    sigma_temp=sigma_L+(market_price-fL)*(sigma_H-sigma_L)/(fH-fL)
    f,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_temp)

    while abs(market_price-f)>error and error_count<100:
        if f<market_price:
            sigma_L=sigma_temp
        else:
            sigma_H=sigma_temp
        
        fL,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_L)
        fH,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_H)
        sigma_temp=sigma_L+(market_price-fL)*(sigma_H-sigma_L)/(fH-fL)
        f,_,_,_,_,_=bs(callput_flag,S,K,T,r,sigma_temp)
        
        error_count+=1

    return sigma_temp


@xloil.func
def gbs_bisection_iv(callput_flag,market_price,S,K,T,r,b,sigma):
    error=0.00001
    error_count=0

    sigma_L=0.001
    sigma_H=10
    
    fL,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_L)
    fH,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_H)
    sigma_temp=sigma_L+(market_price-fL)*(sigma_H-sigma_L)/(fH-fL)
    f,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_temp)

    while abs(market_price-f)>error and error_count<100:
        if f<market_price:
            sigma_L=sigma_temp
        else:
            sigma_H=sigma_temp
        
        fL,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_L)
        fH,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_H)
        sigma_temp=sigma_L+(market_price-fL)*(sigma_H-sigma_L)/(fH-fL)
        f,_,_,_,_,_=gbs(callput_flag,S,K,T,r,b,sigma_temp)
        
        error_count+=1

    return sigma_temp


@xloil.func
def call_bs():
    wb=xw.Book.caller()
    sht=wb.sheets[0]
    
    callput_flag=sht.range('C4').value
    r=sht.range('C8').value
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

    market_price=sht.range('M4').value
    one_S=sht.range('C5').value
    one_K=sht.range('C6').value
    one_T=sht.range('C7').value
    one_sigma=sht.range('C9').value
    newton_iv=bs_newton_iv(callput_flag,market_price,one_S,one_K,one_T,r,one_sigma)
    bisection_iv=bs_bisection_iv(callput_flag,market_price,one_S,one_K,one_T,r,one_sigma)
    sht.range('M5').value=newton_iv
    sht.range('M6').value=bisection_iv

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


@xloil.func
def call_gbs():
    wb=xw.Book.caller()
    sht=wb.sheets[1]
    
    callput_flag=sht.range('C3').value
    S=sht.range('C4').value
    K=sht.range('C5').value
    T=sht.range('C6').value
    r=sht.range('C7').value
    b=sht.range('C8').value
    sigma=sht.range('C9').value
    
    price,delta,gamma,theta,vega,rho=gbs(callput_flag,S,K,T,r,b,sigma)
    sht.range('F3').value=price
    sht.range('F4').value=delta
    sht.range('F5').value=gamma
    sht.range('F6').value=vega
    sht.range('F7').value=theta
    sht.range('F8').value=rho
    
    market_price=sht.range('F9').value
    newton_iv=gbs_newton_iv(callput_flag,market_price,S,K,T,r,b,sigma)
    bisection_iv=gbs_bisection_iv(callput_flag,market_price,S,K,T,r,b,sigma)
    sht.range('F10').value=newton_iv
    sht.range('F11').value=bisection_iv



