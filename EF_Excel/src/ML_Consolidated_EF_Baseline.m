% ------------------------------------------------------------------------------------
% 1.  Set up Portfolio Parameters
% ------------------------------------------------------------------------------------
num_AL=size(AL_list,1);
num_obs=size(hist_data_used,1);

port_mvo=Portfolio;
port_mvo=Portfolio(port_mvo, 'Name','GIS MVO Prototype');
port_mvo=setAssetList(port_mvo, transpose(AL_list));
port_mvo=setInitPort(port_mvo, transpose(curr_wts));
     
ConSet=portcons('PortValue', sum(curr_wts),num_AL, 'AssetLims', transpose(AL_mins), transpose(AL_maxs), 'GroupLims', transpose(group_combinations), transpose(group_mins), transpose(group_maxs),'GroupComparison', transpose(relative_comboA),transpose(relative_mins),transpose(relative_maxs),transpose(relative_comboB));
ConSet_a=ConSet(:,1:end-1);
ConSet_b=ConSet(:,end);
port_mvo=setInequality(port_mvo, ConSet_a, ConSet_b);


% ------------------------------------------------------------------------------------
% 2. Create the Efficient Frontier
% ------------------------------------------------------------------------------------

% 2A - NORMAL RUN
rng(num_seed);
port_mvo=setAssetMoments(port_mvo, base_TR, base_covar);
base_eff_wts=estimateFrontier(port_mvo, num_ef);

% 2B - OPTIONAL RUN
%    VERY SLOW finding weights for evenly targeted risk  - maybe 15 seconds per risk target?
%    base_endpoint_wts = estimateFrontierLimits(port_mvo);
%    [base_endpoint_risk, base_endpoint_return] = estimatePortMoments(port_mvo, base_endpoint_wts);
%    base_target_risks = linspace( base_endpoint_risk (1,1), base_endpoint_risk (2,1), num_ef);
%    base_eff_wts=estimateFrontierByRisk(port_mvo, base_target_risks);

[base_eff_risk, base_eff_return]=estimatePortMoments(port_mvo, base_eff_wts);


% ------------------------------------------------------------------------------------
% 3. Save the Output
% ------------------------------------------------------------------------------------

output_riskT=transpose(base_eff_risk);
output_returnT=transpose(base_eff_return);
output_wts=base_eff_wts;
