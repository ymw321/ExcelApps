

% ------------------------------------------------------------------------------------
% 1.  Set up Portfolio Parameters and Inititalize 
% ------------------------------------------------------------------------------------

num_AL=size(AL_list,1);
num_obs=size(hist_data_used,1);

port_cvar=PortfolioCVaR;
port_cvar=PortfolioCVaR(port_cvar, 'Name','GIS CVaR Prototype');
port_cvar=setAssetList(port_cvar, transpose(AL_list));
port_cvar=setInitPort(port_cvar, transpose(curr_wts));


ConSet=portcons('PortValue', sum(curr_wts),num_AL, 'AssetLims', transpose(AL_mins), transpose(AL_maxs), 'GroupLims', transpose(group_combinations), transpose(group_mins), transpose(group_maxs),'GroupComparison', transpose(relative_comboA),transpose(relative_mins),transpose(relative_maxs),transpose(relative_comboB));
ConSet_a=ConSet(:,1:end-1);
ConSet_b=ConSet(:,end);
port_cvar=setInequality(port_cvar, ConSet_a, ConSet_b);

base_TR_monthly = base_TR/12;
base_covar_monthly = base_covar/12;
cv_scenarios = mvnrnd(base_TR_monthly, base_covar_monthly, num_cv);
port_cvar=PortfolioCVaR(port_cvar, 'Scenarios', cv_scenarios);

port_cvar=PortfolioCVaR(port_cvar, 'ProbabilityLevel', prob_cv);


% ------------------------------------------------------------------------------------
% 2.  Reun the CVAR Optimization
% ------------------------------------------------------------------------------------

wb=waitbar(0,'Calculating CVAR portfolios..');
rng(num_seed);
cv_eff_wts=estimateFrontier(port_cvar, num_ef);

% calculate TR & Vol.  TR can actually be calculated for all rows at once, but volatility is tougher  
for j = 1:num_ef
    cv_eff_return (j) = transpose(cv_eff_wts(:,j))*base_TR;    %  1x50 * 50x1
    cv_eff_risk (j) =sqrt(transpose(cv_eff_wts(:,j))*(base_covar*cv_eff_wts(:,j)));  % 1x50 ( 50x50 * 50x1)
end

delete(wb)


% ------------------------------------------------------------------------------------
% 3.  Save the Output
% ------------------------------------------------------------------------------------

output_riskT = cv_eff_risk;
output_returnT = cv_eff_return;
output_wts = cv_eff_wts;
