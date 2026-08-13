import pandas as pd, numpy as np, json, os, textwrap
from pathlib import Path
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.model_selection import LeaveOneOut, cross_val_predict
from sklearn.inspection import permutation_importance
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

OUT=Path('/mnt/data/lkq_work')
XLS=Path('/mnt/data/Retail Sales Case Study.xlsx')
df=pd.read_excel(XLS, sheet_name=0, engine='openpyxl')
df.columns=['date','sales_m','digital_m','radio_m','tv_k']
df['date']=pd.to_datetime(df['date'])
df['year']=df.date.dt.year
df['quarter']=df.date.dt.quarter
df['qkey']=df.year.astype(str)+' - Q'+df.quarter.astype(str)
gdp=pd.read_excel(XLS, sheet_name=1, engine='openpyxl')
gdp.columns=['qkey','gdp_change']
df=df.merge(gdp,on='qkey',how='left')
df['covid_lockdown']=((df.date>='2020-04-01')&(df.date<='2020-06-30')).astype(int)
df['covid_relaxation']=((df.date>='2020-07-01')&(df.date<='2020-09-30')).astype(int)
df['trend']=np.arange(len(df))
# Convert TV to million dollars for common unit. Keep raw too.
df['tv_m']=df['tv_k']/1000.0

features=['digital_m','radio_m','tv_m','gdp_change','covid_lockdown','covid_relaxation','trend']
X=df[features]
y=df.sales_m
Xc=sm.add_constant(X)
ols=sm.OLS(y,Xc).fit(cov_type='HC3')
# Reduced marketing-only model, because n is small and macro shocks can be collinear
mfeatures=['digital_m','radio_m','tv_m']
mols=sm.OLS(y,sm.add_constant(df[mfeatures])).fit(cov_type='HC3')
# Standardized coefficients for comparable importance
Xs=StandardScaler().fit_transform(X)
ys=(y-y.mean())/y.std(ddof=0)
std_model=sm.OLS(ys, sm.add_constant(Xs)).fit(cov_type='HC3')
std_coefs=pd.Series(np.asarray(std_model.params)[1:],index=features)
std_p=pd.Series(np.asarray(std_model.pvalues)[1:],index=features)
# LOOCV Ridge for predictive sanity check
ridge=Pipeline([('scale',StandardScaler()),('model',Ridge(alpha=10.0))])
pred=cross_val_predict(ridge,X,y,cv=LeaveOneOut())
metrics={'r2_in_sample':float(ols.rsquared),'adj_r2':float(ols.rsquared_adj),'loocv_r2':float(r2_score(y,pred)),'loocv_mae':float(mean_absolute_error(y,pred)),'loocv_rmse':float(mean_squared_error(y,pred)**0.5)}
# vif excluding dummies/trend but full values
vif=pd.Series([variance_inflation_factor(X.values,i) for i in range(X.shape[1])],index=features)
# Correlations
corr=df[['sales_m','digital_m','radio_m','tv_m','gdp_change']].corr()['sales_m'].drop('sales_m')
# Annual and phase summaries
annual=df.groupby('year').agg(sales=('sales_m','sum'),avg_monthly_sales=('sales_m','mean'),digital=('digital_m','sum'),radio=('radio_m','sum'),tv_m=('tv_m','sum')).reset_index()
annual['sales_growth_pct']=annual.sales.pct_change()*100
phases=np.select([df.covid_lockdown.eq(1),df.covid_relaxation.eq(1)],['Lockdown Q2-20','Relaxation Q3-20'],'Other months')
df['phase']=phases
phase=df.groupby('phase').agg(months=('sales_m','size'),avg_sales=('sales_m','mean'),avg_digital=('digital_m','mean'),avg_radio=('radio_m','mean'),avg_tv_m=('tv_m','mean')).reset_index()
quarter=df.groupby(['year','quarter','qkey']).agg(sales=('sales_m','sum'),gdp=('gdp_change','first'),digital=('digital_m','sum'),radio=('radio_m','sum'),tv_m=('tv_m','sum')).reset_index()
# Effects and 95 CI in original units
params=[]
for f in features:
    ci=ols.conf_int().loc[f]
    params.append({'feature':f,'coef':float(ols.params[f]),'p':float(ols.pvalues[f]),'ci_low':float(ci[0]),'ci_high':float(ci[1]),'std_coef':float(std_coefs[f]),'std_p':float(std_p[f])})
params=pd.DataFrame(params)
# marketing-only estimates for interpretation
mparams=[]
for f in mfeatures:
    ci=mols.conf_int().loc[f]
    mparams.append({'feature':f,'coef':float(mols.params[f]),'p':float(mols.pvalues[f]),'ci_low':float(ci[0]),'ci_high':float(ci[1])})
mparams=pd.DataFrame(mparams)
# rolling/time correlations and spend efficiency descriptive
summary={
'n':len(df),'date_min':str(df.date.min().date()),'date_max':str(df.date.max().date()),
'sales_total':float(df.sales_m.sum()),'sales_mean':float(df.sales_m.mean()),
'digital_total':float(df.digital_m.sum()),'radio_total':float(df.radio_m.sum()),'tv_total_m':float(df.tv_m.sum()),
'metrics':metrics,'corr':corr.to_dict(),'vif':vif.to_dict(),
'ols_params':params.to_dict('records'),'marketing_params':mparams.to_dict('records'),
'annual':annual.replace({np.nan:None}).to_dict('records'),'phase':phase.to_dict('records')}
(OUT/'analysis_results.json').write_text(json.dumps(summary,indent=2))
df.to_csv(OUT/'model_input_clean.csv',index=False)
params.to_csv(OUT/'model_coefficients.csv',index=False)

# styling
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.titlesize':15,'axes.labelsize':10})
blue='#3B82F6'; navy='#0B1F3A'; teal='#14B8A6'; orange='#F59E0B'; red='#EF4444'; gray='#64748B'; light='#E2E8F0'
# 1 monthly sales + covid
fig,ax=plt.subplots(figsize=(11,4.3))
ax.plot(df.date,df.sales_m,color=blue,lw=2.5,marker='o',ms=4)
ax.axvspan(pd.Timestamp('2020-04-01'),pd.Timestamp('2020-06-30'),color=red,alpha=.12,label='Lockdown Q2 2020')
ax.axvspan(pd.Timestamp('2020-07-01'),pd.Timestamp('2020-09-30'),color=orange,alpha=.12,label='Relaxation Q3 2020')
ax.set_title('Monthly sales were volatile; no clean structural collapse during lockdown',loc='left',color=navy,fontweight='bold')
ax.set_ylabel('Sales, $M'); ax.grid(axis='y',alpha=.2); ax.spines[['top','right']].set_visible(False); ax.legend(frameon=False,ncol=2,loc='upper right')
fig.tight_layout(); fig.savefig(OUT/'monthly_sales.png',dpi=180,transparent=False); plt.close(fig)
# quarter sales/gdp dual axis
fig,ax=plt.subplots(figsize=(10.5,4.2)); x=np.arange(len(quarter))
ax.bar(x,quarter.sales,color=blue,alpha=.86,label='Quarterly sales')
ax.set_ylabel('Sales, $M'); ax.set_xticks(x); ax.set_xticklabels(quarter.qkey.str.replace(' - ', '\n'),fontsize=8)
ax2=ax.twinx(); ax2.plot(x,quarter.gdp*100,color=orange,marker='o',lw=2.3,label='GDP change')
ax2.axhline(0,color=gray,lw=.8); ax2.set_ylabel('GDP change, %')
ax.set_title('GDP shock and sales do not move one-for-one',loc='left',color=navy,fontweight='bold')
ax.spines['top'].set_visible(False); ax2.spines['top'].set_visible(False); ax.grid(axis='y',alpha=.18)
lines,labels=ax.get_legend_handles_labels(); l2,lab2=ax2.get_legend_handles_labels(); ax.legend(lines+l2,labels+lab2,frameon=False,ncol=2,loc='upper right')
fig.tight_layout(); fig.savefig(OUT/'quarter_gdp_sales.png',dpi=180); plt.close(fig)
# standardized coefficient plot
plot=params.sort_values('std_coef')
labels={'digital_m':'Digital','radio_m':'Radio','tv_m':'TV','gdp_change':'GDP','covid_lockdown':'COVID lockdown','covid_relaxation':'COVID relaxation','trend':'Time trend'}
fig,ax=plt.subplots(figsize=(8.8,4.8)); colors=[teal if v>0 else red for v in plot.std_coef]
ax.barh([labels[x] for x in plot.feature],plot.std_coef,color=colors)
ax.axvline(0,color=navy,lw=1); ax.set_xlabel('Standardized coefficient')
ax.set_title('Digital is the strongest positive signal after controls',loc='left',color=navy,fontweight='bold')
ax.grid(axis='x',alpha=.2); ax.spines[['top','right','left']].set_visible(False)
for i,(v,p) in enumerate(zip(plot.std_coef,plot.std_p)):
    ax.text(v+(0.02 if v>=0 else -0.02),i,f"p={p:.2f}",va='center',ha='left' if v>=0 else 'right',fontsize=8,color=gray)
fig.tight_layout(); fig.savefig(OUT/'standardized_effects.png',dpi=180); plt.close(fig)
# channel scatter small multiples
fig,axs=plt.subplots(1,3,figsize=(11,3.65))
for ax,col,title,c in zip(axs,['digital_m','radio_m','tv_k'],['Digital ($M)','Radio ($M)','TV ($K)'],[blue,teal,orange]):
    ax.scatter(df[col],y,s=35,color=c,alpha=.75,edgecolor='white',linewidth=.4)
    z=np.polyfit(df[col],y,1); xx=np.linspace(df[col].min(),df[col].max(),100); ax.plot(xx,np.polyval(z,xx),color=navy,lw=1.8)
    r=np.corrcoef(df[col],y)[0,1]; ax.set_title(f'{title}\nr = {r:.2f}',fontweight='bold',color=navy); ax.set_xlabel('Spend'); ax.grid(alpha=.18); ax.spines[['top','right']].set_visible(False)
axs[0].set_ylabel('Sales, $M'); fig.suptitle('Only Digital shows a strong, consistent linear relationship',x=.02,ha='left',color=navy,fontweight='bold',fontsize=15)
fig.tight_layout(rect=[0,.01,1,.91]); fig.savefig(OUT/'channel_scatter.png',dpi=180); plt.close(fig)
# annual bars
fig,ax=plt.subplots(figsize=(8,4));
ax.bar(annual.year.astype(str),annual.sales,color=[gray,blue,teal]);
for i,v in enumerate(annual.sales): ax.text(i,v+25,f'${v:,.0f}M',ha='center',fontweight='bold',color=navy)
ax.set_title('Annual sales improved in 2021, then softened in 2022',loc='left',color=navy,fontweight='bold'); ax.set_ylabel('Annual sales, $M'); ax.grid(axis='y',alpha=.2); ax.spines[['top','right']].set_visible(False)
fig.tight_layout(); fig.savefig(OUT/'annual_sales.png',dpi=180); plt.close(fig)
print(json.dumps(summary,indent=2))
