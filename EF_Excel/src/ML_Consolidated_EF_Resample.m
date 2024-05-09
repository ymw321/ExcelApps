

% ------------------------------------------------------------------------------------
% 1.  Set up Portfolio Parameters and Inititalize 
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

rs_sav_mean(1,num_AL)=0;
rs_sav_vol(1,num_AL)=0;
rs_sav_cov(num_AL,num_AL)=0;
rs_sav_fullmatrix(1,1:7+num_AL)=0;
rs_sav_buckets(1:num_ef,1:7+num_AL)=0;

% ------------------------------------------------------------------------------------
% 2.  Get Min & Max Risk Portfolios for baseline assumptions to establish bucket size
% ------------------------------------------------------------------------------------

port_mvo=setAssetMoments(port_mvo, base_TR, base_covar);
base_endpoint_wts = estimateFrontierLimits(port_mvo);
[base_endpoint_risk, base_endpoint_return] = estimatePortMoments(port_mvo, base_endpoint_wts);
rs_low_risk = base_endpoint_risk(1,1);
rs_high_risk = base_endpoint_risk(2,1);
rs_bucket = (rs_high_risk - rs_low_risk) / num_ef;


% ------------------------------------------------------------------------------------
% 3.  Set up Bootstrapping Data (if necessary) 
% ------------------------------------------------------------------------------------

rs_bs_weights = 0;
hist_data_bootstrap = 0;

% 3a.  Calculate bootstrap data size
if rs_type (1,1:4) == 'Boot'
   hist_data_num_rows=size(hist_data_used,1);
   hist_data_num_indices=size(hist_data_used,2);
   hist_data_num_covar=hist_data_num_indices*hist_data_num_indices;
end

% 3b.  Calculate bootstrap weight vector (if rel wt = 1, uniform)
if rs_type (1,1:4) == 'Boot'
   rs_bs_factor = rs_relwt ^ (4/(num_obs-1));
   for i = 1:num_obs
       rs_bs_weights (i) = rs_bs_factor ^ -((i-1)/4);
   end    
end

% 3c.  If bootstrapping with expected return assumptions, need to convert
%      historical data to standard normals to an adjusted-expected data
%      Adj Data:  = ( Historical - Avg(hist) * Vol(exp)/Vol(hist) + Avg(exp) 						
%      Need to adjust annual expected data to quarterly basis first

if rs_type (1,1:4) == 'Boot' & return_source (1,1:4) == 'Hist'
   hist_data_bootstrap = hist_data_used;
end   

if rs_type (1,1:4) == 'Boot' & return_source (1,1:4) == 'Expe' 
   hist_qtrly_avg = mean (hist_data_used);
   hist_qtrly_std = std (hist_data_used);
   base_qtrly_avg = transpose(base_TR)/4;
   base_qtrly_std = transpose(sqrt(diag(base_covar)))/2;
   hist_data_bootstrap = hist_data_used;
   hist_data_bootstrap = bsxfun(@minus, hist_data_bootstrap, hist_qtrly_avg);
   hist_data_bootstrap = bsxfun(@rdivide, hist_data_bootstrap, hist_qtrly_std);
   hist_data_bootstrap = bsxfun(@times, hist_data_bootstrap, base_qtrly_std);
   hist_data_bootstrap = bsxfun(@plus, hist_data_bootstrap, base_qtrly_avg);
end


% ------------------------------------------------------------------------------------
% 4.  create resampled assumptions and run efficient frontier one by one
% ------------------------------------------------------------------------------------

wb=waitbar(0,'Resampling Scenarios..');
rng(num_seed);
for i = 1:num_rs

  if rs_type (1,1:4) == 'MVNo'
  %  create a multi variate normal scenario using base annual TR & Covariance
     rs_scenario = mvnrnd(base_TR, base_covar, num_obs);
     rs_mean = mean(rs_scenario);                            % 1 x num_indices
     rs_vol = sqrt ( var(rs_scenario));                      % 1 x num_indices   
     rs_cov = cov(rs_scenario);                              % num_indices x num_indices
     xx=1
  else
  %  create a bootstrapped scenario using quarterly "observations", 
  %  adjust scenario to annual
     hist_data_bs_stats = bootstrp(1, @(x)[mean(x) reshape(cov(x),1,hist_data_num_covar)], hist_data_bootstrap, 'weights', rs_bs_weights);
     hist_data_bs_mean = hist_data_bs_stats (1, 1:hist_data_num_indices)
     hist_data_bs_cov = reshape(hist_data_bs_stats (1, hist_data_num_indices+1:end),hist_data_num_indices,hist_data_num_indices)
     rs_mean = hist_data_bs_mean * 4;                        % 1 x num_indices           
     rs_cov = hist_data_bs_cov * 4;                          % num_indices x num_indices    
     rs_vol = transpose(sqrt(diag(rs_cov)));                 % 1 x num_indices
  end
       
  %  save highlights of the scenario
  rs_sav_mean(i,:)=rs_mean;
  rs_sav_vol(i,:)=rs_vol;
  rs_sav_corr(:,:,i)=corrcov(rs_cov);
  
  %  create an efficient portfolio for the resampled scenario
  port_mvo=setAssetMoments(port_mvo, rs_mean, rs_cov);
  rs_eff_wts=estimateFrontier(port_mvo, num_ef);
  [rs_eff_risk, rs_eff_return]=estimatePortMoments(port_mvo, rs_eff_wts);
  
  %  save all the efficient portfolios in the next "number_porfolios" rows
  %  col_1 = loop#, col_2 = frontier#, col_3 = bucket
  %  col_4 = base_TR, col_5 = base_vol, col 6 = scen_TR, col_7 = scen _vol
  %  col 8 - end = scen_eff_wts
  %  sr = start row for matrix
  sr = (i-1)*num_ef;   
  rs_sav_fullmatrix(sr+1:sr+num_ef,:)=0;
  rs_sav_fullmatrix(sr+1:sr+num_ef,1)=i;
  rs_sav_fullmatrix(sr+1:sr+num_ef,2)=1:num_ef;
  rs_sav_fullmatrix(sr+1:sr+num_ef,6)=rs_eff_return;
  rs_sav_fullmatrix(sr+1:sr+num_ef,7)=rs_eff_risk;
  rs_sav_fullmatrix(sr+1:sr+num_ef,8:num_AL+8-1)=transpose(rs_eff_wts);

  waitbar(i/num_rs,wb,sprintf('%u',i))
 
end

delete(wb)


% ------------------------------------------------------------------------------------
% 4. Add data to the resampled portfolio full matrix  
%    one row for each frontier point * scenario combination
% ------------------------------------------------------------------------------------

%  rs_sav_fullmatrix(:,4)=rs_sav_fullmatrix(:,8:num_AL+8-1)*base_TR;
%  does not work:  rs_sav_fullmatrix(:,5)=sqrt(rs_sav_fullmatrix(:,8:num_AL+8-1)*(base_covar*transpose(rs_sav_fullmatrix(:,8:num_AL+8-1))));
  
%  fullmatrix_size = size (rs_sav_fullmatrix);
%  fullmatrix_rows = fullmatrix_size(1,1);
%  for j = 1:fullmatrix_rows  

  for j = 1:num_rs*num_ef
    
    % calculate TR & Vol.  TR can actually be calculated for all rows at once, but volatility is tougher  
    rs_sav_fullmatrix(j,4)=rs_sav_fullmatrix(j,8:num_AL+8-1)*base_TR;
    rs_sav_fullmatrix(j,5)=sqrt(rs_sav_fullmatrix(j,8:num_AL+8-1)*(base_covar*transpose(rs_sav_fullmatrix(j,8:num_AL+8-1))));
    
    % calculate bucket number
    rs_sav_fullmatrix(j,3)=min(round((rs_sav_fullmatrix(j,5)- rs_low_risk)/rs_bucket +0.5,0),num_ef);

    % accumulate data in the summary buckets
    sum_row = rs_sav_fullmatrix(j,3);
    rs_sav_buckets(sum_row,2)=rs_sav_buckets(sum_row,2)+1;  
    rs_sav_buckets(sum_row,8:num_AL+8-1)=rs_sav_buckets(sum_row,8:num_AL+8-1)+rs_sav_fullmatrix(j,8:num_AL+8-1);
    
  end


% ------------------------------------------------------------------------------------
% 4. Summarize data to the resampled portfolio summary matrix
%    one row for each frontier point 
% ------------------------------------------------------------------------------------

for k = 1:num_ef
    %  create summary data for each bucket
    %  col_1 = bucket#, col_2 = num_rs, col_3 = not used
    %  col_4 = base_TR, col_5 = base_vol, col 6 = not used, col_7 = not used
    %  col 8 - end = scen_eff_wtss
     
    rs_sav_buckets(k,1)=k;
    if rs_sav_buckets(k,2) > 0
       %  calculate average weight for bucket "k"
       rs_sav_buckets(k,8:num_AL+8-1)=rs_sav_buckets(k,8:num_AL+8-1)/rs_sav_buckets(k,2);
       %  recalculate return and volatility for those average wt
       rs_sav_buckets(k,4)=rs_sav_buckets(k,8:num_AL+8-1)*base_TR;
       rs_sav_buckets(k,5)=sqrt(rs_sav_buckets(k,8:num_AL+8-1)*(base_covar*transpose(rs_sav_buckets(k,8:num_AL+8-1))));
    end
    
end

% ------------------------------------------------------------------------------------
% 5. Save the Output
% ------------------------------------------------------------------------------------

       output_riskT = transpose(rs_sav_buckets(:,5));
       output_returnT = transpose(rs_sav_buckets(:,4));
       output_wts = transpose(rs_sav_buckets(:,8:end));
       
       save rs_data.mat rs_sav_mean rs_sav_vol rs_sav_corr
