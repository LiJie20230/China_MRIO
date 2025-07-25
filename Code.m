clear all

path='./Trade';
cd(path);

load Export_31 

Export=Export_Final3(:,2:end);

load Import_31 

Import=Export_Final3(:,2:end);

clearvars Export_Final3

cd .. % go back to the previous dictionary

%%%%%2018 Output and VA
N_Province=Name_EX(2:end,2);
for i=1:size(N_Province,1)
     
    VA(:,i)=xlsread('OUTPUT&VA',N_Province{i},'G6:G47');
    Output(:,i)=xlsread('OUTPUT&VA',N_Province{i},'K6:K47');
    
end
%%%%preliminary treatment
Check_ov=Output-VA;
Check_ov(find(Check_ov>0))=0;

%%%%%2017 Output 
[~,~,N_Province]=xlsread('Name','name');
for i=1:size(N_Province,1)-1

  IO2017{i}=xlsread(['.\province\',N_Province{i+1,4}],'Sheet1');

end

%%%%% 2017 Trade Matrix
%%% Check re-export
for i=1:size(N_Province,1)-1
  
   IT=IO2017{i};
   IT(isnan(IT))=0;
   
   Check_re=IT(1:42,52)+IT(1:42,53)-IT(1:42,58);
   Check_re_ratioEX=IT(1:42,52)./(IT(1:42,52)+IT(1:42,53));
   Check_re_ratioOL=IT(1:42,53)./(IT(1:42,52)+IT(1:42,53));
   Check_re_ratioEX(isnan(Check_re_ratioEX))=0;
   Check_re_ratioOL(isnan(Check_re_ratioOL))=0;
        
   Check_re_ratioIM=IT(1:42,55)./(IT(1:42,55)+IT(1:42,56));
   Check_re_ratioIL=IT(1:42,56)./(IT(1:42,55)+IT(1:42,56));
   Check_re_ratioIM(isnan(Check_re_ratioIM))=0;
   Check_re_ratioIL(isnan(Check_re_ratioIL))=0;
   
   for j=1:42
       
   if Check_re(j,:)<0
      Check_re(j,:)=0;
   end
   
   end
   
   Trade(1:42,i)=IT(1:42,52)-Check_re_ratioEX.*Check_re;%%Export   
   Trade(43:84,i)=IT(1:42,55)-Check_re_ratioIM.*Check_re;%%Import
           
   Trade1(1:42,i)=IT(1:42,53)-Check_re_ratioOL.*Check_re;%%Outflow 
   Trade1(43:84,i)=IT(1:42,56)-Check_re_ratioIL.*Check_re;%%Inflow
      
     %%%Constraint for cross entropy
    IT(1:42,52)=Trade(1:42,i);%%%export replacement
    IT(1:42,55)=Trade(43:84,i);%%%import replacement
    IT(1:42,53)=Trade1(1:42,i);%%%outflow replacement
    IT(1:42,56)=Trade1(43:84,i);%%%inflow replacement 
           
    EX_Service_2017(:,i)=IT(28:42,52);
    IM_Service_2017(:,i)=IT(28:42,55);
    
    Z_2017{i}=IT(1:42,1:42);
    F_2017{i}=[IT(1:42,44:45),IT(1:42,47),IT(1:42,49:50)];
        
    
    Local_Supply_2017(:,i)=IT(1:42,58)-IT(1:42,52)-IT(1:42,53);
    Domestic_Supply_2017(:,i)=IT(1:42,53);
    Domestic_Demand_2017(:,i)=IT(1:42,56);
     
    Intermediate_2017(:,i)=IT(1:42,43);  
    
    Total_Demand_20171(:,i)=IT(1:42,43)+IT(1:42,48)+IT(1:42,51);  
   
   P2018{i}=IT;
end

Local_Supply_2017(find(Local_Supply_2017<0))=0;

Local_Supply_2017=Local_Supply_2017';
Domestic_Supply_2017=Domestic_Supply_2017';
Domestic_Demand_2017=Domestic_Demand_2017';

Total_Supply_2017=Local_Supply_2017+Domestic_Supply_2017;
Total_Demand_2017=Local_Supply_2017+Domestic_Demand_2017;

Total_Supply_2017(find(Total_Supply_2017<0.1))=0;

Proportion_2017=Total_Demand_2017./sum(Total_Demand_2017,1);

PP_2=Intermediate_2017./Total_Demand_20171;

%%%%%Supply-Demand Matrix 2017
for i=1:42

SDM1_2017{i}=[Local_Supply_2017(:,i),Domestic_Supply_2017(:,i)]./sum(sum([Local_Supply_2017(:,i),Domestic_Supply_2017(:,i)]));
SDM2_2017{i}=[Local_Supply_2017(:,i),Domestic_Demand_2017(:,i)]./sum(sum([Local_Supply_2017(:,i),Domestic_Demand_2017(:,i)]));

end

for i=1:42
  Priori_2017{i}=[SDM1_2017{i},SDM2_2017{i}];
  
  %Priori_2017{i}=ones(31,4);
end

%%%% 2018 China
NIOT_2018=xlsread('2018IOT','D9:BG57');
Bridge_139t42=xlsread('Name','Bridge Sector');
NIOT_2018=NIOT_2018*Bridge_139t42;

Export_NA=NIOT_2018(1:42,52);
Import_NA=NIOT_2018(1:42,54);

National_Demand=NIOT_2018(1:42,43)+NIOT_2018(1:42,48)+NIOT_2018(1:42,51);
National_supply=NIOT_2018(1:42,56)-Export_NA;
%Demand_Domestic_2018=Proportion_2017.*National_Demand';
%xlswrite('Demand_Domestic_2018',Demand_Domestic_2018);

%%%%Output-Va 2018
%%%% Output scaled by national one
Output_2=NIOT_2018(1:42,56).*(Output./sum(Output,2));

%%%% VA scaled by national one
VA_2=NIOT_2018(48,1:42)'.*(VA./sum(VA,2));


Check2=Output_2-VA_2;
Check2(find((Output_2-VA_2)>0))=0;

Output_2=Output_2-Check2;

%%%%%% total demand
for i=1:31
    A_proportion1= xlsread('A_proportion.xlsx',i);%%%% Replace A with 2018 Chongqing Zhejiang Guangdong Hunan Anhui
    A_proportion1(isnan(A_proportion1))=0;
    A_proportion1(isinf(A_proportion1))=0;
    PP_2018(:,i)=sum(A_proportion1.*transpose(Output_2(:,i)),2)  ; 
    A_proportion2{i}=A_proportion1;
end

Total_Demand_p=PP_2018./PP_2;

Total_Demand=Total_Demand_p./sum(Total_Demand_p,2).*National_Demand;

%Total_Demand_C=(PP_2018./sum(PP_2018,2)).*National_Demand;%% for sector 1:27

%Total_Demand_S=(Proportion_2017.*National_Demand')';%% for Construction

%Total_Demand_O=(Proportion_2017.*National_Demand')'

%Total_Demand=[Total_Demand_C(1:27,:);Total_Demand_S(28,:);Total_Demand_O(29:42,:)];

%%%%%% RAS TRADE 2018
Trade_Total=xlsread('Name','Total Trade');
Trade_scaled=(Trade_Total./sum(Trade_Total,1)).*[sum(Import_NA(1:27,1),1),sum(Export_NA(1:27,1),1)];%%%% IM & EX
Trade_scaled=Trade_scaled';
%%% Commodity
for i=1:31
     IOP=P2018{i};
     Export_A(1:42,i)=IOP(1:42,52);
     Import_A(1:42,i)=IOP(1:42,55);
end
%Export
Export_Commodity=RAS(Export_NA(1:27),Trade_scaled(2,:),Export_A(1:27,1:end));
%Import
Import_Commodity=RAS(Import_NA(1:27),Trade_scaled(1,:),Import_A(1:27,1:end));

%%% Service
%%%% Using 2017 foreign trade structure
Trade_S=[sum(IM_Service_2017,1);sum(EX_Service_2017,1)]';
Trade_scaled=(Trade_S./sum(Trade_S,1)).*[sum(Import_NA(28:42,1),1),sum(Export_NA(28:42,1),1)];%%%% IMS & EXS
Trade_scaled=Trade_scaled';
%Export
Export_Service=RAS(Export_NA(28:42),Trade_scaled(2,:),EX_Service_2017);
%Import
Import_Service=RAS(Import_NA(28:42),Trade_scaled(1,:),IM_Service_2017);

%%%%%%%% Foregin Trade for 31 provinces
Export_31=[Export_Commodity;Export_Service];
Import_31=[Import_Commodity;Import_Service];

%%%%%%Check Output-Export 
lo=Output_2-Export_31;
lo(find(lo>0))=0;

Export_2=Export_31+lo;
Import_2=Import_31+lo;
Import_2(find(Import_2<0))=0;

%%%%%%%%%% Cross Entropy 
Lom=Total_Demand-Import_2;
Lom(find(Lom>0))=0;
Export_3=Export_2+Lom;
Import_3=Import_2+Lom;

Export_3(find(Export_3<0))=0;

Row_Cons=(Output_2-Export_3)';

Total_Demand1=Total_Demand-Import_3;

Demand_Domestic_2018=(Total_Demand1)';
Demand_Domestic_2018(find(Demand_Domestic_2018<0))=0;

Demand_Domestic_20181=(Demand_Domestic_2018./sum(Demand_Domestic_2018,1)).*sum(Row_Cons,1);
%Error=Row_Cons2-Row_Cons;


%%%%%%%%%% Trade estimate
for i=1:42
   OldFolder = pwd;    
   Trade_matrix_2018{i}=Gams_2018(Row_Cons(:,i),Demand_Domestic_20181(:,i),Priori_2017{i});
   cd(OldFolder);
end

for i=1:42
     Trade_matrix_Sec=Trade_matrix_2018{i};
  for j=1:31
    Trade_Sec(i,(j-1)*4+1:(j-1)*4+4)=Trade_matrix_Sec(j,:);
  end
end

%%%%%%% additional constraint 
AD_CON=sum(VA_2,1)-(sum(Export_3,1)-sum(Import_3,1));

%%%%%%
for i=1:31
   Trade_Sec1(:,(i-1)*2+1:(i-1)*2+2)=[Trade_Sec(:,(i-1)*4+2),Trade_Sec(:,(i-1)*4+4)];
end

for i=1:31
    XXX=sum(Trade_Sec1,1);
    XXX1(i,1)=XXX(:,(i-1)*2+1);
    XXX1(i,2)=XXX(:,(i-1)*2+2);
    
    delta(:,i)=XXX(:,(i-1)*2+1)-XXX(:,(i-1)*2+2);
end

F_Province_2018=AD_CON-delta;

%%%%% Preliminary Distribution 2017-2018
path='.\Proinvce IO';
cd(path);

[~,~,Name_G]=xlsread('Name_G','Name','C1:C31');

for i=1:31
    Name_G1{i}=char(Name_G{i});
    
    if  size(Name_G1{i},2)~=1;
   
        Name_G2=xlsread(Name_G1{i},'T','D7:BI55');
        Name_G2(isnan(Name_G2))=0;
      else   
    end
    F_2018=[Name_G2(1:42,44:45),Name_G2(1:42,47),Name_G2(1:42,49:50)];
    P_F{i}= F_2018./sum(sum(F_2018));
end

[~,~,br]=xlsread('Name_G','Br','B2:AF32');

path='C:\Users\CEADS HR\OneDrive - University College London\IO TABLE\China IO Compliation\2#Table\2018 MRIO\2018 MRIO\Construction';
cd(path);

for i=1:31
    for j=1:31
    br1(i,j)=br{i,j};
    
    end
end
br1(isnan(br1))=0;

for i=1:31
    if isempty(find(br1(:,i)==1))==1;
    br2(i)=0;
    else
    br2(i)=find(br1(:,i)==1);
    end
end

count=0;
for i=1:31
   F_proportion=F_2017{i}./sum(sum(F_2017{i},1),2);
    
   if i==br2
       count=count+1;
        F_proportion=P_F{count};
       
   end
    F_Province_2018P{i}=F_proportion.*F_Province_2018(i);    
   
end



for i=1:31
   Z_Province_2018P{i}=(A_proportion2{i}.*transpose(Output_2(:,i))); 
    
   Z_4_RAS{i}=[Z_Province_2018P{i},F_Province_2018P{i}];
end

%%%%% Matrix for Cross Entropy £¨with negetive sign£©GRAS
for i=1:31
    

     Z_4_RAS1=Z_4_RAS{i};
    Z_4_RAS1(isnan(Z_4_RAS1))=0;
        
    Row_Constraint{i}=transpose(Row_Cons(i,:))+Import_3(:,i)-(Trade_Sec1(:,(i-1)*2+1)-Trade_Sec1(:,(i-1)*2+2));%%-transpose(Error(i,:)
    Column_Constraint{i}=[transpose(Output_2(:,i)-VA_2(:,i)),sum(F_Province_2018P{i},1)];
    
    RowC=Row_Constraint{i};
    ColumnC=Column_Constraint{i};
    RowC(find(RowC<0.01))=0;
    ColumnC(find(ColumnC<0.01&ColumnC>0))=0;
    
    OldFolder = pwd;  
    check(:,i)=sum(RowC,1)-sum(ColumnC,2);
    SRIO_raw{i}=Optimization(Z_4_RAS1,RowC,ColumnC); 
   
    cd(OldFolder);
end

%%%%%% Write SRIO table for 31 2018
[~,~,IO_Layout]=xlsread('Name','IO layout');
[~,~,Namenn]=xlsread('Name','N_Province','A1:A31');

for i=1:31
    SRIO_raw1=SRIO_raw{i};
    
    SRIO_20181{i}= [SRIO_raw1,Export_3(:,i),Trade_Sec1(:,(i-1)*2+1),Import_3(:,i),Trade_Sec1(:,(i-1)*2+2),repmat(0,42,1),Output_2(:,i);transpose(VA_2(:,i)),0,0,0,0,0,0,0,0,0,0,0;transpose(Output_2(:,i)),0,0,0,0,0,0,0,0,0,0,0];
    
   % xlswrite('SRIO2018.xlsx', IO_Layout,Namenn{i});
    %xlswrite('SRIO2018.xlsx', SRIO_20181{i},Namenn{i},'D11:BD54');
    %%%Check
end


%%%Check
for i=1:31
    SRIO_raw2=SRIO_20181{i};
    
    Check_Table(:,i)=sum(SRIO_raw2(1:42,1:49),2)-sum(SRIO_raw2(1:42,50:51),2)-SRIO_raw2(1:42,53);
   
end

Check_Table(find(Check_Table<0.1))=0;


clearvars -except SRIO_20181 Trade_matrix_2018

%%% MRIO CONSTRUCTION
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%% Using transforamtion matrix from national table
xlsread('201801')


%%% Type A to Type B transformations reexport included
for i=1:31
    
    Z_In=SRIO_20181{i};
    Ratio_FM=Z_In(1:42,50)./(sum(Z_In(1:42,1:47),2)+Z_In(1:42,49));%%%%%%%% For foreign trade tolerant re-export only in domestic trade
    Ratio_PM=Z_In(1:42,51)./(sum(Z_In(1:42,1:47),2)+Z_In(1:42,49));%%% For domestic trade tolerant re-export only in domestic trade
     
    Ratio_PM(isnan(Ratio_PM))=0;
    Ratio_PM(isinf(Ratio_PM))=0;
    Ratio_FM(isnan(Ratio_FM))=0;
    Ratio_FM(isinf(Ratio_FM))=0;
    
    Ratio_D(:,i)=1-Ratio_FM- Ratio_PM;
    Ratio_D1=Ratio_D(:,i);
    Ratio_D1(find(Ratio_D1>0))=0; 
    
    Ratio_D(find(Ratio_D<0))=0; 
    
    Ratio_PM=Ratio_PM+Ratio_D1;
    Ratio_PM1=Ratio_PM;
    Ratio_PM1(find(Ratio_PM1>0))=0;
    Ratio_PM(find(Ratio_PM<0))=0; 
    
    Ratio_FM1=Ratio_FM+Ratio_PM1;  
       
    Outflow_re_f(:,i)=Ratio_FM1.*Z_In(1:42,49);
    Outflow_re_d(:,i)=Ratio_PM.*Z_In(1:42,49);
    
end

Outflow_re=Outflow_re_f+Outflow_re_d;
Outflow_re1=Outflow_re(:);

%%%Sample data preparation%%%%%%%
[~,~,y]=xlsread('Sample Data','Commodity name','A2:A13');
name1=char(y);

for i=1:length(y)
  space=name1(i,:);  
  space=strrep(space,' ','');
  Trade_T=xlsread('Sample Data',space,'B3:AF33');
  Trade{i}=Trade_T; 
  Trade2{i}=Trade_T';
end
 
td=zeros(31);
for i=1:length(y)
  td=cell2mat(Trade(i));
  td(isnan(td))=0;
  %for k=1:31
  %    td(k,k)=0;      
  %end
  td=td';
  td=td(:);
  td1=log(td);
  td1(isnan(td1))=0;
  td1(isinf(td1))=0;
  Trade{i}=td1;
end
%%%% Trade2
td=zeros(31);
for i=1:length(y)
  td=cell2mat(Trade2(i));
  td(isnan(td))=0;
  for k=1:31
      td(k,k)=0;      
  end
  td=td(:);
  td1=log(td);
  td1(isnan(td1))=0;
  td1(isinf(td1))=0;
  Trade2{i}=td1;
end

%%%%% distance %%%%%
d=xlsread('Sample Data','Distance','B2:AF32');
d1=d(:);

%%%%% Total supply and Total demand
for i=1:31
    tra=SRIO_20181{i};
    TFM_old((i-1)*42+1:(i-1)*42+42,1)=tra(1:42,49);%%% Outflow
    TFM_old((i-1)*42+1:(i-1)*42+42,2)=tra(1:42,51); %%% Inflow
end
TFM_old(find(TFM_old<0.1))=0;
TFM_old(isnan(TFM_old))=0;

Real_trade=TFM_old(:,2)-(Outflow_re1);

%%%% estimate Re_outflow 
Outflow_re1(find(Real_trade<0 ))=0;

Real_trade_1=TFM_old-Outflow_re1;%%%% Re_outflow
Real_trade_1(find(Real_trade_1<0))=0;
%%%%%%%% new domestic trade %%%%%%

for i=1:31
    
    Z_In=SRIO_20181{i};
    Z_In(1:42,49)=Real_trade_1((i-1)*42+1:(i-1)*42+42,1);
    
    Z_In(1:42,51)=Real_trade_1((i-1)*42+1:(i-1)*42+42,2);
    SRIO_20182{i}=Z_In;
end

    %%% Type A to Type B transformations reexport included
for i=1:31
    
    Z_In=SRIO_20182{i};
    Ratio_FM=Z_In(1:42,50)./sum(Z_In(1:42,1:47),2);%%%%%%%% For foreign trade tolerant re-export only in domestic trade
    Ratio_PM=Z_In(1:42,51)./sum(Z_In(1:42,1:47),2);%%% For domestic trade tolerant re-export only in domestic trade
        
    
    Ratio_PM(isnan(Ratio_PM))=0;
    Ratio_PM(isinf(Ratio_PM))=0;
    Ratio_FM(isnan(Ratio_FM))=0;
    Ratio_FM(isinf(Ratio_FM))=0;
      
    
    Ratio_D(:,i)=1-Ratio_FM- Ratio_PM;
    Ratio_D1=Ratio_D(:,i);
    Ratio_D1(find(Ratio_D1>0))=0; 
    
    Ratio_D(find(Ratio_D<0))=0; 
    
    Ratio_PM=Ratio_PM+Ratio_D1;
    Ratio_PM1=Ratio_PM;
    Ratio_PM1(find(Ratio_PM1>0))=0;
    Ratio_PM(find(Ratio_PM<0))=0; 
    
    Ratio_FM1=Ratio_FM+Ratio_PM1;  
       
    Z_D=Ratio_D(:,i).*Z_In(1:42,1:47);
    Z_C=Ratio_PM.*Z_In(1:42,1:47);
    Z_D(43,1:47)=sum(Ratio_FM1.*Z_In(1:42,1:47),1);
     
    Z_D1{i}=Z_D;
    Z_CM{i}=Z_C;
        
end

%%%%% New TFM
%%%%% Total supply and Total demand
for i=1:31
    tra=SRIO_20182{i};
    TFM((i-1)*42+1:(i-1)*42+42,1)=tra(1:42,49);%%% Outflow
    TFM((i-1)*42+1:(i-1)*42+42,2)=tra(1:42,51); %%% Inflow
end
TFM(find(TFM<0.1))=0;
TFM(isnan(TFM))=0;


TFM1=TFM; %% FOR REGRESSION

%%%%% demand and supply for regression 11 commodity%%%%
Coal_1=zeros(31*31,24);

for i=1:31
      Coal(i,1:2) = TFM1((i-1)*42+2,1:2);%% Coal
      Coal(i,3:4) = TFM1((i-1)*42+11,1:2);%%%Coke
      Coal(i,5:6) = TFM1((i-1)*42+3,1:2);%%%Petro
      Coal(i,7:8) = TFM1((i-1)*42+14,1:2);%%%Steel
      Coal(i,9:10) = TFM1((i-1)*42+4,1:2);%%%Metal
      Coal(i,11:12) = TFM1((i-1)*42+5,1:2);%%%Non-Metal
      Coal(i,13:14) = TFM1((i-1)*42+13,1:2);%%% Mineral Construction Material
      Coal(i,15:16) = TFM1((i-1)*42+9,1:2);%%% Timber
      Coal(i,17:18) = TFM1((i-1)*42+12,1:2);%%% Fertilizer
      Coal(i,19:20) = TFM1((i-1)*42+1,1:2);%%% Food
      Coal(i,21:22) = TFM1((i-1)*42+7,1:2);%%% Cotton
      Coal(i,23:24)= TFM1((i-1)*42+25,1:2); %%% Electricity
end

 for n=1:12 
      Coal_1(:,(n-1)*2+2)=repmat(Coal(1:31,(n-1)*2+2),31,1);  
  for j=1:31
      Coal_1((j-1)*31+1:(j-1)*31+31,(n-1)*2+1)=Coal(j,(n-1)*2+1);
  end  
 end


%%%% Regression for 12 commodity%%%%

for i=1:12
   X_supply(:,i)=log(Coal_1(:,(i-1)*2+1));
   X_demand(:,i)=log(Coal_1(:,(i-1)*2+2));
end
x3=log(d1);

X_supply(isnan(X_supply))=0;
X_demand(isnan(X_demand))=0;
x3(isnan(x3))=0;

X_supply(isinf(X_supply))=0;
X_demand(isinf(X_demand))=0;
x3(isinf(x3))=0;

for i=1:12
  X=[ones(961,1),X_supply(:,i),X_demand(:,i),x3];
  [b(i,:),bint,r,rint,stats(:,i)]=regress(Trade{i},X);
  
  for n=1:4
   
   [h,p{n,i},ci,stats_ttest{n,i}]=ttest(Trade{i},X(:,n));
  
   end
end

%%% OD-Raw%%%
Bridge=xlsread('Sample Data','com-bridge');
Bridge(isnan(Bridge))=0;


% construction corresponding relationship
%%%% Commodity Sector
for i=1:31
   
      Sort(i,1:2) = TFM((i-1)*42+2,1:2);%% Coal
      Sort(i,3:4) = TFM((i-1)*42+11,1:2);%%%Coke 
      Sort(i,5:6) = TFM((i-1)*42+3,1:2);%%%Petro extracting, 
      
      Sort(i,7:8)= TFM((i-1)*42+14,1:2);   %%%Steel
      Sort(i,9:10) = TFM((i-1)*42+15,1:2); %%%Steel
      Sort(i,11:12) = TFM((i-1)*42+16,1:2);%%%Steel
      Sort(i,13:14) = TFM((i-1)*42+17,1:2);%%%Steel
      Sort(i,15:16) = TFM((i-1)*42+18,1:2);%%%Steel
      Sort(i,17:18) = TFM((i-1)*42+19,1:2);%%%Steel
      Sort(i,19:20) = TFM((i-1)*42+20,1:2);%%%Steel
      Sort(i,21:22) = TFM((i-1)*42+21,1:2);%%%Steel
      Sort(i,23:24) = TFM((i-1)*42+22,1:2);%%%Steel
      Sort(i,25:26) = TFM((i-1)*42+23,1:2);%%%Steel
      Sort(i,27:28) = TFM((i-1)*42+24,1:2);%%%Steel
      
      
      Sort(i,29:30) = TFM((i-1)*42+4,1:2);%%%Metal
      Sort(i,31:32) = TFM((i-1)*42+5,1:2);%%%Non-Metal
      Sort(i,33:34) = TFM((i-1)*42+13,1:2);%%% Mineral Construction Material
      
    
      Sort(i,35:36) = TFM((i-1)*42+9,1:2);  %%% Timber
      Sort(i,37:38) = TFM((i-1)*42+10,1:2);%%% Timber paper
      
      Sort(i,39:40) = TFM((i-1)*42+12,1:2);%%% Fertilizer
      
      Sort(i,41:42) = TFM((i-1)*42+1,1:2);%%% Food
      Sort(i,43:44) = TFM((i-1)*42+6,1:2);%%% Food
      
      Sort(i,45:46) = TFM((i-1)*42+7,1:2);%%% Cotton
      Sort(i,47:48) = TFM((i-1)*42+8,1:2);%%% Cotton
      Sort(i,49:50) = TFM((i-1)*42+25,1:2);%%% E
end
%%%%%% Tertiary Sector
% Maxumi entropy
for i=1:31
   Sort_s(i,1:2) = TFM((i-1)*42+26,1:2);%%% E
   Sort_s(i,3:4) = TFM((i-1)*42+27,1:2);%%% E
   Sort_s(i,5:6) = TFM((i-1)*42+28,1:2);%%% E
   Sort_s(i,7:8) = TFM((i-1)*42+29,1:2);%%% E
   Sort_s(i,9:10) = TFM((i-1)*42+30,1:2);%%% E
   Sort_s(i,11:12) = TFM((i-1)*42+31,1:2);%%% E
   Sort_s(i,13:14) = TFM((i-1)*42+32,1:2);%%% E
   Sort_s(i,15:16) = TFM((i-1)*42+33,1:2);%%% E
   Sort_s(i,17:18) = TFM((i-1)*42+34,1:2);%%% E
   Sort_s(i,19:20) = TFM((i-1)*42+35,1:2);%%% E
   Sort_s(i,21:22) = TFM((i-1)*42+36,1:2);%%% E
   Sort_s(i,23:24) = TFM((i-1)*42+37,1:2);%%% E
   Sort_s(i,25:26) = TFM((i-1)*42+38,1:2);%%% E
   Sort_s(i,27:28) = TFM((i-1)*42+39,1:2);%%% E
   Sort_s(i,29:30) = TFM((i-1)*42+40,1:2);%%% E
   Sort_s(i,31:32) = TFM((i-1)*42+41,1:2);%%% E
   Sort_s(i,33:34) = TFM((i-1)*42+42,1:2);%%% E
end


%%%%%
Beta(1:2,:)=b(1:2,:);
Beta(3,:)=b(3,:);
Beta(4:14,:)=repmat(b(4,:),11,1);
Beta(15,:)=b(5,:);
Beta(16,:)=b(6,:);
Beta(17,:)=b(7,:);
Beta(18:19,:)=repmat(b(8,:),2,1);
Beta(20,:)=b(9,:);
Beta(21:22,:)=repmat(b(10,:),2,1);
Beta(23:24,:)=repmat(b(11,:),2,1);
Beta(25,:)=b(12,:);

Beta1=Beta';

for i=1:25
 for j=1:31

  Xt1((j-1)*31+1:(j-1)*31+31,i)=repmat(Sort(j,(i-1)*2+1),31,1);
 
 end
end


%%%%%%%
for i=1:25
    
    ABC=log(Xt1(:,i));
    ABC(isnan(ABC))=0;
    ABC(isinf(ABC))=0;
     
     
    ABC2=log(repmat(Sort(:,(i-1)*2+2),31,1));
    ABC2(isnan(ABC2))=0;
    ABC2(isinf(ABC2))=0;
    Sector(:,i) = [ones(961,1),ABC,ABC2,x3]*Beta1(:,i);

end

Sector(find(Sector<0))=0.00000000000001;
Sector(isinf(Sector))=0;
Sector(isnan(Sector))=0;

%%%% Certain sectors Replace by    
            %Sector(:,1)=Trade2{1}; %%% coal 
            %Sector(:,2)=Trade2{2}; %%% coke 
            %Sector(:,3)=Trade2{3}; %%% petro 
            %Sector(:,4)=Trade2{4}; %%% steel
            %Sector(:,15)=Trade2{5}; %%% metal mine
            %Sector(:,16)=Trade2{6}; %%% non metal mine
            %Sector(:,17)=Trade2{7}; %%% non furrous 
            %Sector(:,18)=Trade2{8}; %%% timeber
            %Sector(:,20)=Trade2{9}; %%% fertiliser
            %Sector(:,21)=Trade2{10}; %%% agriculture 
            %Sector(:,23)=Trade2{11}; %%% cotton 



        for i=1:25
            for j=1:31

                Sector_OD((i-1)*31+1:(i-1)*31+31,j)= Sector((j-1)*31+1:(j-1)*31+31,i) ;

            end
        end
        
        [~,~,Sector_name]=xlsread('Sample Data','Commodity name','B14:B38');
        Sector_name1=char(Sector_name);

        [~,~,Frame]=xlsread('Sample Data','Frame');


        for i=1:25
         
         space=Sector_name1(i,:);
         space=strrep(space,' ','');

         ODX=Sector_OD((i-1)*31+1:(i-1)*31+31,1:31);
         ODX=ODX';
         for k=1:31
           ODX(k,k)=0;
         end
         %xlswrite('Sector.xlsx',Frame,space);
         %xlswrite('Sector.xlsx',ODX,space,'B2:AD30');
         %xlswrite('Sector.xlsx',space(:,:),space,'A1:A1');
         ODX1{i}=ODX;
        end


%%%%% RAS %%%%
for i=1:25
    
    ODX=ODX1{i};
    
    U=Sort(:,(i-1)*2+1); %% Row Constraint
    V=transpose(Sort(:,(i-1)*2+2)); %% Conlumn Constraint
    
    U(isnan( U))=0;
    V(isnan(V))=0;

    diff1=1;
    diff2=1;
    eta=1;
    diss=1;

  while  ( sum(diff2,1) >0.1 || sum(diff1,2)>0.1 )
        
% R
RR=U./sum(ODX,2);

ODX=diag(RR)*ODX;

ODX(isnan(ODX))=0;

% S
SS=V./sum(ODX,1);

ODX=ODX*diag(SS);

ODX(isnan(ODX))=0;

diff1= abs((sum(ODX,1)-V))./V ;
diff1(isnan(diff1))=0;
diff2= abs((sum(ODX,2)-U))./U; 
diff2(isnan(diff2))=0;


  eta=eta+1;

    end
   
 ODX2{i}=ODX;

end

%%%% Service Sector 25 to 42
%%%%Entropy maxi

for i=1:17
    OldFolder = pwd;    
    ODX3{i}=Doubly2_EX(Sort_s(:,(i-1)*2+1:(i-1)*2+2));
    cd(OldFolder);
end

for i=1:17
 ODX2{25+i}=ODX3{i};
end

%%%%% Bridge sector from Gravity model 
Bridge_GM=xlsread('Name','Bridge Sector');
Bridge_GM(isnan(Bridge_GM))=0;

for i=1:42
   
   BB= Bridge_GM(:,i);
   ODX4{i}=ODX2{find(BB==1)};
    
end

n=31;

%%%%%%% Interregional Trade flow estimate over
%%%%derive the RPC
for i=1:42
   RPC((i-1)*n+1:(i-1)*n+n,1:n)=ODX4{i}./sum(ODX4{i},1);
end

RPC(isnan(RPC))=0;
RPC(isinf(RPC))=0;

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%% Construct Intermediate AND Final Demand Matrix
for i=1:n
    
    
    Z_INT=Z_CM{i};
    for j=1:42
       
        Z_INT2((j-1)*n+1:(j-1)*n+n,(i-1)*47+1:(i-1)*47+47)=Z_INT(j,:).*RPC((j-1)*n+1:(j-1)*n+n,i);
       
    end
   
end
    
    
%%% M is SxC / CxS
%%%%convert to city-sector to city-sector
%%%%by row
for i=1:42
for   j=1:n
       Z_INT3((j-1)*42+i,:)=Z_INT2((i-1)*n+j,:);
end
end
Z_INT3(isnan(Z_INT3))=0;

%%%%%Domestic Matrix
for i=1:n
    Z_D2=Z_D1{i};
    Z_INT3((i-1)*42+1:(i-1)*42+42,(i-1)*47+1:(i-1)*47+47)=Z_D2(1:42,1:47);
end

%%% Change layout from 42X47 to 42X42 + 42X5
for i=1:n
    for j=1:n
    Z_INT4((i-1)*42+1:(i-1)*42+42,(j-1)*42+1:(j-1)*42+42)=Z_INT3((i-1)*42+1:(i-1)*42+42,(j-1)*47+1:(j-1)*47+42);
    F_IN4((i-1)*42+1:(i-1)*42+42,(j-1)*5+1:(j-1)*5+5)=Z_INT3((i-1)*42+1:(i-1)*42+42,(j-1)*47+43:(j-1)*47+47);
    end
end
Table_raw=[Z_INT4,F_IN4];

%%% add Export and outflow 
%outflow and %export
for i=1:n
Outf=SRIO_20181{i};
Outf=[Outf(1:42,48),Outf(1:42,53)];
Table_raw((i-1)*42+1:(i-1)*42+42,n*42+n*5+1:n*42+n*5+2)=Outf;
end

%%%Import and Inflow
for i=1:n
IMMM=Z_D1{i};
IMMM1=IMMM(43,1:42);
IMF=IMMM(43,43:47);
Table_raw(42*n+1,(i-1)*42+1:(i-1)*42+42)=IMMM1;
Table_raw(42*n+1,(i-1)*5+n*42+1:(i-1)*5+n*42+5)=IMF;
end


for i=1:n
    IOT=SRIO_20181{i};
    
    Table_raw(n*42+2,(i-1)*42+1:(i-1)*42+42)=IOT(43,1:42);%%%% VA
    
    Table_raw(n*42+3,(i-1)*42+1:(i-1)*42+42)=IOT(44,1:42);%%%% Oputput
end

xlswrite('MRIO20183.xlsx',Table_raw);





