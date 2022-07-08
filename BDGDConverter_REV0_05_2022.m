clc
clear all

%Leitura das tabelas
%-------------------------------------------//
COD_ID = [0;1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40;41;42;43;44;45;46;47;48;49;50;51;52;53;54;55;56;57;58;59;60;61;62;63;64;65;66;67;68;69;70;71;72;73;74;75;76;77;78;79;80;81;82;83;84;85;86;87;88;89;90;91;92;93;94;95;96;97;98;99;100];
TEM = [0;110;115;120;121;125;127;208;216;216.500000000000;220;230;231;240;254;380;400;440;480;500;600;750;1000;2200;3200;3600;3785;3800;3848;3985;4160;4200;4207;4368;4560;5000;6000;6600;6930;7960;8670;11400;11900;12000;12600;12700;13200;13337;13530;13800;13860;14140;14190;14400;14835;15000;15200;19053;19919;21000;21500;22000;23000;23100;23827;24000;24200;25000;25800;27000;30000;33000;34500;36000;38000;40000;44000;45000;45400;48000;60000;66000;69000;72500;88000;88200;92000;100000;120000;121000;123000;131600;131630;131635;138000;145000;230000;345000;500000;750000;1000000];
DESCR={'Não informado';'110 V';'115 V';'120 V';'121 V';'125 V';'127 V';'208 V';'216 V';'216.5 V';'220 V';'230 V';'231 V';'240 V';'254 V';'380 V';'400 V';'440 V';'480 V';'500 V';'600 V';'750 V';'1 kV';'2.2 kV';'3.2 kV';'3.6 kV';'3.785 kV';'3.8 kV';'3.848 kV';'3.985 kV';'4.16 kV';'4.2 kV';'4.207 kV';'4.368 kV';'4.56 kV';'5 kV';'6 kV';'6.6 kV';'6.93 kV';'7.96 kV';'8.67 kV';'11.4 kV';'11.9 kV';'12 kV';'12.6 kV';'12.7 kV';'13.2 kV';'13.337 kV';'13.53 kV';'13.8 kV';'13.86 kV';'14.14 kV';'14.19 kV';'14.4 kV';'14.835 kV';'15 kV';'15.2 kV';'19.053 kV';'19.919 kV';'21 kV';'21.5 kV';'22 kV';'23 kV';'23.1 kV';'23.827 kV';'24 kV';'24.2 kV';'25 kV';'25.8 kV';'27 kV';'30 kV';'33 kV';'34.5 kV';'36 kV';'38 kV';'40 kV';'44 kV';'45 kV';'45.4 kV';'48 kV';'60 kV';'66 kV';'69 kV';'72.5 kV';'88 kV';'88.2 kV';'92 kV';'100 kV';'120 kV';'121 kV';'123 kV';'131.6 kV';'131.63 kV';'131.635 kV';'138 kV';'145 kV';'230 kV';'345 kV';'500 kV';'750 kV';'1000 kV'};
TTEN = table(COD_ID,TEM,DESCR);
clear COD_ID TEM DESCR

CTMT = readtable('CTMT.xlsx');
SSDMT = readtable('SSDMT.xlsx');
SSDBT = readtable('SSDBT.xlsx');
SEGCON = readtable('SEGCON.xlsx');
EQTRD = readtable('EQTRD.xlsx');
UNTRD = readtable('UNTRD.xlsx');
UCBT = readtable('UCBT.xlsx');
UCMT = readtable('UCMT.xlsx');
PIP = readtable('PIP.xlsx');
%-------------------------------------------//

% //Dados de entrada
%-------------------------------------------//
prompt= 'Alimentador (BDGD) COD_ID da camada (CTMT): ';
Ali=input(prompt,'s');

freq=60;

prompt= 'Potência de curto-circuito trifásica (MVA): ';
mvasc3=input(prompt,'s');

prompt= 'Potência de curto-circuito monofásica (MVA): ';
mvasc1=input(prompt,'s');

prompt= 'Análise-> Dias úteis(UT), Sábados(SAB) ou Domingos(DOM): ';
Dia_analise=input(prompt,'s');
%-------------------------------------------//

% //Equivalente Thevenin/Circuit
% //----------------------------------------//
SUB = CTMT(strcmp(CTMT.COD_ID, Ali), :);
DESCR = TTEN(TTEN.COD_ID == str2double(SUB{1,7}), 3);
DESCR{1,1} = regexprep(DESCR{1, 1}, 'V$', '');
DESCR{1,1} = regexprep(DESCR{1, 1}, 'k$', '');
DESCR{1,1} = regexprep(DESCR{1, 1}, ' ', '');
DESCR=str2double(DESCR{1,1});

SSDBT = SSDBT(strcmp(SSDBT.CTMT, Ali), :);
SSDMT = SSDMT(strcmp(SSDMT.CTMT, Ali), :);

SSDMT_CONTAINS=contains(SSDMT.PN_CON_1,SSDMT.PN_CON_2);
SSDMT_BUS = SSDMT(SSDMT_CONTAINS(:,1) == 0,:);
bus1 = string(SSDMT_BUS{1,3});
% //----------------------------------------//

% //LineCodes
% //----------------------------------------//
SSDMT_13 = unique(SSDMT(:,13),'stable');
SEGCON_MT=SEGCON(ismember(SEGCON.COD_ID, SSDMT_13.TIP_CND), :);
rows_MT=size(SEGCON_MT,1);
i=1;

SSDBT_14 = unique(SSDBT(:,14),'stable');
SEGCON_BT=SEGCON(ismember(SEGCON.COD_ID, SSDBT_14.TIP_CND), :);
rows_BT=size(SEGCON_BT,1);
% //----------------------------------------//

% //Lines
% //----------------------------------------//
rows_SSDMT=size(SSDMT,1);
rows_SSDBT=size(SSDBT,1);
Code_first_line = string(SSDMT_BUS{1,2});
% //----------------------------------------//
% //Trafos
% //----------------------------------------//
UNTRD = UNTRD(strcmp(UNTRD.CTMT, Ali), :);
rows_UNTRD=size(UNTRD,1);
SSDBT_FIL=unique(SSDBT(:,5),'stable');
PN_2_UNTRD=table;
base_kv_ten_lin=UNTRD{1,14};
% //----------------------------------------//

% //----------------------------------------//
% //Loads/LoadsShapes
% //----------------------------------------//
UCBT = UCBT(strcmp(UCBT.CTMT, Ali), :);
UCMT = UCMT(strcmp(UCMT.CTMT, Ali), :);
PIP = PIP(strcmp(PIP.CTMT, Ali), :);
rows_UCBT=size(UCBT,1);
rows_UCMT=size(UCMT,1);
rows_PIP=size(PIP,1);
kwh=table;

R_UTEIS = "0.875718121, 0.7379304, 0.659207807, 0.633476025, 0.61442096, 0.678764605, 0.804322151, 0.925609108, 0.812108228, 0.738965351, 0.841846193, 0.950237106, 0.976420214, 0.851260093, 0.761985332, 0.77603361, 0.837538914, 1.369080181, 2.392437728, 1.778169959, 1.383523896, 1.353454872, 1.166310245, 1.081178899";
C_UTEIS = "0.540290381, 0.522840237, 0.500970213, 0.490532996, 0.508517788, 0.547920027, 0.571156332, 0.702277961, 1.102078597, 1.403910599, 1.577798928, 1.549089466, 1.283974133, 1.426076057, 1.579967453, 1.589447048, 1.55488613, 1.453487725, 1.202275416, 1.006271074, 0.859995139, 0.770251199, 0.670245313, 0.585739789";
IN_UTEIS = "0.371480548, 0.336803009, 0.342382469, 0.356678593, 0.422471297, 0.550753421, 0.654840865, 1.110008527, 1.596005663, 1.705449919, 1.852359559, 1.730589546, 1.008336908, 1.532508343, 1.870899751, 1.799217941, 1.690571957, 1.463406208, 0.925729658 0.739420845, 0.559664929, 0.520173023, 0.45983599, 0.40041103";
SP_UTEIS = "0.715469742, 0.667620058, 0.66107404, 0.682765098, 0.652029173, 0.698830552, 0.861324277, 1.01744213, 1.137088058, 1.209694169, 1.250450996, 1.263345256, 1.250983732, 1.204525696, 1.187173553, 1.171167595, 1.187824991, 1.180237526, 1.160738967, 1.084776851, 1.040818453, 0.998823054, 0.915379224, 0.80041681";
A4_UTEIS = "0.797062444, 0.776931084, 0.770858997, 0.788911111, 0.826906519, 0.91039443, 0.995234812, 1.154512925, 1.264221535, 1.292436805, 1.326093804, 1.295605964, 1.182864062, 1.318314054, 1.347710884, 1.338456486, 1.309773825, 1.158663378, 0.740455996, 0.53623314, 0.508044577, 0.65548448, 0.869557771, 0.835270918";
IP_UTEIS = "2.013976041, 2.013976041, 2.013976041, 2.013976041, 2.013976041, 2.013976041, 1.594394572, 0.167832588, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.251748882, 1.846143453, 2.013976041, 2.013976041, 2.013976041, 2.013976041";

R_SAB = "0.9479247 0.778990851 0.644163368 0.622945212 0.599629318 0.631423811 0.624300097 0.784340665 0.875381538 0.921509082 0.976392549 1.073391253 1.020256255 1.095033215 1.052651324 1.11871709 1.173223331 1.241993938 1.359572292 1.487065095 1.47764092 1.301080174 1.138071683 1.054302238";
C_SAB = "0.745844961 0.695560012 0.660555887 0.656953401 0.6675735 0.706258203 0.697273586 0.741337801 0.997543038 1.270040175 1.46779961 1.474424886 1.360264397 1.25427073 1.249090243 1.196283151 1.18359598 1.041567259 1.094021382 1.114351578 1.048219812 0.969091752 0.908246631 0.799832026";
IN_SAB = "0.719701223 0.678856369 0.65797677 0.688729903 0.740285843 0.824971236 0.974840711 1.249931292 1.558711053 1.626715041 1.578187507 1.487839002 1.16260595 1.127922538 1.195249562 1.166267331 1.106884403 1.007627862 0.937038061 0.882564052 0.738361397 0.649525851 0.643271554 0.59593549";
SP_SAB = "0.764124928 0.705296903 0.701858516 0.663284763 0.66956166 0.689300838 0.818684876 0.945400647 1.050293212 1.174812297 1.197044164 1.205997552 1.209776334 1.206965474 1.202857298 1.168455913 1.178516186 1.146120392 1.165970522 1.14098204 1.091501285 1.052795445 0.971168566 0.879230189";
A4_SAB = "1.051824318 1.027558766 1.022662475 1.033062567 1.032821183 1.073151932 1.104095753 1.149256352 1.189917618 1.206903066 1.194106858 1.141280143 1.049368708 1.030409138 1.020566211 0.996503977 0.957821268 0.909606485 0.88206375 0.851650023 0.817218272 0.785744757 0.750715532 0.721690847";

R_DOM = "0.902446549 0.793209114 0.762242413 0.634578965 0.612704677 0.619183942 0.61238264 0.686929606 0.828068989 1.037885774 1.243109979 1.336588186 1.179317195 1.068303868 1.00666421 0.974474563 1.055412957 1.188204305 1.328681513 1.366134223 1.275224649 1.357337019 1.114345767 1.016568896";
C_DOM = "0.993897432 0.972326705 0.936434735 0.892515887 0.874446892 0.895711227 0.907327765 0.891246254 0.94811259 1.031648282 1.129049038 1.125414381 1.050651267 1.034211027 1.014083065 1.005260628 0.946590527 0.978046403 1.087096631 1.155707182 1.096008472 1.061153529 1.006011016 0.967049066";
IN_DOM = "0.864017127 0.839475836 0.816636209 0.868310784 0.899163682 0.914429973 0.919112742 0.982321722 1.066884737 1.03857701 1.102481975 1.062662815 1.02392571 1.027317706 1.028976738 1.026373181 1.067446416 1.0949689 1.180691398 1.097615015 0.988605309 0.969014647 1.086182031 1.034808338";
SP_DOM = "0.792237795 0.770612686 0.758954136 0.771332012 0.712239535 0.741435859 0.917681788 0.938596341 1.014057683 1.134204229 1.178114981 1.191946067 1.170631599 1.197045856 1.156067675 1.153229582 1.172034243 1.144135328 1.164647755 1.085417461 1.046097267 1.002677709 0.981975564 0.804626847";
A4_DOM = "0.966397294 0.953181849 0.942987579 0.932101485 0.940800001 0.925168058 0.92901207 0.954305205 0.981923834 1.00508849 1.026816252 1.015473749 0.978544231 0.985969777 0.99257738 0.981376001 0.98515142 0.99085031 1.017522936 1.016404235 1.003221898 1.04150091 1.176262898 1.257362137";

if Dia_analise == "UT"
    fid = fopen( 'LoadShapes.txt', 'wt' );
    fprintf(fid, 'New loadshape.Residencial npts = 24 interval = 1 mult = (%s) \n', R_UTEIS);
    fprintf(fid, 'New loadshape.Comercial npts = 24 interval = 1 mult = (%s) \n', C_UTEIS);   
    fprintf(fid, 'New loadshape.Industrial npts = 24 interval = 1 mult = (%s) \n', IN_UTEIS);
    fprintf(fid, 'New loadshape.Serv_Pub npts = 24 interval = 1 mult = (%s) \n', SP_UTEIS);
    fprintf(fid, 'New loadshape.A4 npts = 24 interval = 1 mult = (%s) \n', A4_UTEIS);
    fprintf(fid, 'New loadshape.IP npts = 24 interval = 1 mult = (%s) \n', IP_UTEIS);     
    fclose(fid);
elseif Dia_analise == "SAB"
    fid = fopen( 'LoadShapes.txt', 'wt' );
    fprintf(fid, 'New loadshape.Residencial npts = 24 interval = 1 mult = (%s) \n', R_SAB);
    fprintf(fid, 'New loadshape.Comercial npts = 24 interval = 1 mult = (%s) \n', C_SAB);   
    fprintf(fid, 'New loadshape.Industrial npts = 24 interval = 1 mult = (%s) \n', IN_SAB);
    fprintf(fid, 'New loadshape.Serv_Pub npts = 24 interval = 1 mult = (%s) \n', SP_SAB);
    fprintf(fid, 'New loadshape.A4 npts = 24 interval = 1 mult = (%s) \n', A4_SAB);
    fprintf(fid, 'New loadshape.IP npts = 24 interval = 1 mult = (%s) \n', IP_UTEIS);  
    fclose(fid);
elseif Dia_analise == "DOM"
    fid = fopen( 'LoadShapes.txt', 'wt' );
    fprintf(fid, 'New loadshape.Residencial npts = 24 interval = 1 mult = (%s) \n', R_DOM);
    fprintf(fid, 'New loadshape.Comercial npts = 24 interval = 1 mult = (%s) \n', C_DOM);   
    fprintf(fid, 'New loadshape.Industrial npts = 24 interval = 1 mult = (%s) \n', IN_DOM);
    fprintf(fid, 'New loadshape.Serv_Pub npts = 24 interval = 1 mult = (%s) \n', SP_DOM);
    fprintf(fid, 'New loadshape.A4 npts = 24 interval = 1 mult = (%s) \n', A4_DOM);
    fprintf(fid, 'New loadshape.IP npts = 24 interval = 1 mult = (%s) \n', IP_DOM);  
    fclose(fid);
else
    
end
% //----------------------------------------//

% //Escrever TXTs
% //----------------------------------------//

N1 = sprintf("LineCode_SSDBT_%s",Ali);
N2 = sprintf("LineCode_SSDMT_%s",Ali);
N3 = sprintf("Lines_SSDMT_%s",Ali);
N4 = sprintf("Lines_SSDBT_%s",Ali);
N5 = sprintf("LoadShapes");
N6 = sprintf("Loads_UCBT_%s",Ali);
N7 = sprintf("Loads_UCMT_%s",Ali);
N8 = sprintf("Transformers_%s",Ali);
N9 = sprintf("PIP_%s",Ali);

fid = fopen( 'MainCode.txt', 'wt' );
fprintf(fid, 'clear\n\n');
fprintf(fid, 'New Circuit.%s bus1 = %s basekv = %g pu = 1.00 phases = 3 frequency = %g mvasc3 = %s mvasc1 = %s\n\n', Ali, bus1, DESCR, freq, mvasc3, mvasc1);
fprintf(fid, 'Redirect %s.txt\n',N1);
fprintf(fid, 'Redirect %s.txt\n',N2);
fprintf(fid, 'Redirect %s.txt\n',N3);
fprintf(fid, 'Redirect %s.txt\n',N4);
fprintf(fid, 'Redirect %s.txt\n',N5);
fprintf(fid, 'Redirect %s.txt\n',N6);
fprintf(fid, 'Redirect %s.txt\n',N7);
fprintf(fid, 'Redirect %s.txt\n',N8);
fprintf(fid, 'Redirect %s.txt\n',N9);
fprintf(fid, '\nSet VoltageBases = [%g, %g] \nCalcVoltageBases\n', DESCR, base_kv_ten_lin);
fprintf(fid, '\nNew energymeter.medidor element = line.%s terminal = 1\n', Code_first_line);
fprintf(fid, '\nset tolerance = 0.0000001\nset Maxiter=500\nset mode = daily\nset stepsize = 1h\nset number = 24\nsolve');
fclose(fid);

fid = fopen(sprintf('%s.txt', N2) ,'wt');
for i = 1:rows_MT %Escrita linecode MT
    x= string(SEGCON_MT{i,2});
    c= SEGCON_MT{i,18};
    y= SEGCON_MT{i,15};
    z= SEGCON_MT{i,16};
    fprintf(fid,'New linecode.%s nphases = 3 basefreq = %g units = km Normamps = %g \n~ R1 = %g !ohm/km \n~ X1 = %g !ohm/km \n~ C1 = 0.00\n\n', x, freq, c, y, z);
    if i == rows_MT
        clear x c y z i
    end    
end
fclose(fid);

fid = fopen(sprintf('%s.txt', N1) ,'wt');
for i = 1:rows_BT %Escrita linecode BT
    x= string(SEGCON_BT{i,2});
    
    A = str2double(SEGCON_BT{i,37});
    B = str2double(SEGCON_BT{i,38});
    C = str2double(SEGCON_BT{i,39});
    N = str2double(SEGCON_BT{i,40});
    l=0;
    if A ~= 0  
        l=1;
    end
    if B ~= 0  
        l=l+1;
    end
    if C ~= 0  
        l=l+1;
    end
    if N ~= 0
        l=l+1;
    end
    c= SEGCON_BT{i,18};
    y= SEGCON_BT{i,15};
    z= SEGCON_BT{i,16};
    fprintf(fid,'New linecode.%s nphases = %g basefreq = %g units = km Normamps = %g \n~ R1 = %g !ohm/km \n~ X1 = %g !ohm/km \n~ C1 = 0.00\n\n', x, l, freq, c, y, z);
    if i == rows_BT
        clear x c y z i l A B C N
    end    
end 
fclose(fid);

fid = fopen(sprintf('%s.txt', N3) ,'wt');
for i = 1:rows_SSDMT %Escrita Lines MT
    x= string(SSDMT{i,2});
    c= string(SSDMT{i,3});
    y= string(SSDMT{i,4});
    z= string(SSDMT{i,19});
    l= string(SSDMT{i,13});
    fprintf(fid,'New line.%s phases = 3 bus1 = %s bus2 = %s \n~ length = %s units = m \n~ linecode = %s\n\n', x, c, y, z, l);
    if i == rows_SSDMT
        clear x c y z i l
    end    
end
fclose(fid);

fid = fopen(sprintf('%s.txt', N8) ,'wt');
for i = 1:rows_UNTRD %Escrita TRAFOS
    x= string(UNTRD{i,2}); % COLETA CODIGO/nome TRAFO
    PAC_2 = string(UNTRD{i,5});
    EQTRD_FIL=EQTRD(strcmp(EQTRD.PAC_2,PAC_2),:);
    xhl=EQTRD_FIL{1,32}; % xhl do trafo da iteração
    PAC_1=string(EQTRD_FIL{1,4}); %coleta PAC_1 do trafo
    SSDMT_PN_2=SSDMT(strcmp(SSDMT.PAC_2, PAC_1), :); %filtra SSDMT referente a conexão com o trafo
    isempty(SSDMT_PN_2);
    
    if isempty(SSDMT_PN_2) == 1
        SSDMT_PN_1=SSDMT(strcmp(SSDMT.PAC_1, PAC_1), :);
        PN_2=string(SSDMT_PN_1{1,3}); %coleta o codigo da barra (saida SSDMT e entrada trafo)
    else
        PN_2=string(SSDMT_PN_2{1,4}); %coleta o codigo da barra (saida SSDMT e entrada trafo)
    end
    
    kva=UNTRD{i,20};
    fases=string(UNTRD{i,7});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" || fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABCN"
        conn1 ="wye";
    elseif fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABC"
        conn1 ="delta";
    else
        continue
    end
    res=EQTRD_FIL{1,31};
    kv=UNTRD{i,14};
    
    fases=string(UNTRD{i,8});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" || fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABCN"
        conn2 ="wye";
    elseif fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABC"
        conn2 ="delta";
    else
        continue
    end
    
    tap = UNTRD{i,17};
    noloadloss = UNTRD{i,21}/(kva*10);
    loadloss = UNTRD{i,22}/(kva*10)-noloadloss;
    
    if fases == "A"
        conex = ".1";
    elseif fases == "B"
        conex = ".2";
    elseif fases == "C"
        conex = ".3";
    elseif fases == "AN"
        conex = ".1.0";
    elseif fases == "BN"
        conex = ".2.0";
    elseif fases == "CN"
        conex = ".3.0";
    elseif fases == "AB"
        conex = ".1.2";
    elseif fases == "BC"
        conex = ".2.3";
    elseif fases == "CA"
        conex = ".3.1";
    elseif fases == "ABN"
        conex = ".1.2.0";
    elseif fases == "BCN"
        conex = ".2.3.0";
    elseif fases == "CAN"
        conex = ".3.1.0";
    elseif fases == "ABC"
        conex = ".1.2.3";
    elseif fases == "ABCN"
        conex = ".1.2.3.0";
    end
    PN_2_UNTRD(i,:) = {PN_2};
    fprintf(fid,'New transformer.%s xhl = %g windings = 2 %%loadloss = %g %%noloadloss = %g  \n~ wdg = 1 bus = %s kv = %g kva = %g conn = %s %%r = %g tap = 0.9565 \n~ wdg = 2 bus = BT_%s%s kv = %g kva = %g conn = %s %%r = %g tap = %g\n\n', x, xhl, loadloss, noloadloss, PN_2, DESCR, kva, conn1, res, PN_2, conex, kv, kva, conn2, res, tap);
    if i == rows_UNTRD
        clear x c y z i l fases a PAC_1 xhl PAC_2 kv conn2 kva PN_2
    end    
end  
fclose(fid);

fid = fopen(sprintf('%s.txt', N4) ,'wt');
for i = 1:rows_SSDBT %Escrita Lines BT
    x= string(SSDBT{i,2});
   
    fases=string(SSDBT{i,10});
    if fases == "A" || fases == "B" || fases == "C" || fases == "N"
            phase = "1";
    elseif  fases == "AN" || fases == "BN" || fases == "CN"    
            phase = "2";
    elseif  fases == "AB" || fases == "BC" || fases == "CA"
            phase = "2";
    elseif  fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABC"
            phase = "3"';   
    elseif  fases == "ABCN"
            phase = "4"; 
    end    
    
    c= string(SSDBT{i,3});
    y= string(SSDBT{i,4});   
    
    if fases == "A"
        conex2 = ".1";
    elseif fases == "B"
        conex2 = ".2";
    elseif fases == "C"
        conex2 = ".3";
    elseif fases == "N"
        conex2 = ".4";        
    elseif fases == "AN"
        conex2 = ".1.4";
    elseif fases == "BN"
        conex2 = ".2.4";
    elseif fases == "CN"
        conex2 = ".3.4";
    elseif fases == "AB"
        conex2 = ".1.2";
    elseif fases == "BC"
        conex2 = ".2.3";
    elseif fases == "CA"
        conex2 = ".3.1";
    elseif fases == "ABN"
        conex2 = ".1.2.4";
    elseif fases == "BCN"
        conex2 = ".2.3.4";
    elseif fases == "CAN"
        conex2 = ".3.1.4";
    elseif fases == "ABC"
        conex2 = ".1.2.3";
    elseif fases == "ABCN"
        conex2 = ".1.2.3.4";
    end
    PN_2_CONTAIN= find(strcmp(PN_2_UNTRD.Var1,c));
    
    if PN_2_CONTAIN ~= '0' 
        if fases == "A"
            conex1 = ".1";
        elseif fases == "B"
            conex1 = ".2";
        elseif fases == "C"
            conex1 = ".3";
        elseif fases == "N"
            conex1 = ".4";            
        elseif fases == "AN"
            conex1 = ".1.0";
        elseif fases == "BN"
            conex1 = ".2.0";
        elseif fases == "CN"
            conex1 = ".3.0";
        elseif fases == "AB"
            conex1 = ".1.2";
        elseif fases == "BC"
            conex1 = ".2.3";
        elseif fases == "CA"
            conex1 = ".3.1";
        elseif fases == "ABN"
            conex1 = ".1.2.0";
        elseif fases == "BCN"
            conex1 = ".2.3.0";
        elseif fases == "CAN"
            conex1 = ".3.1.0";
        elseif fases == "ABC"
            conex1 = ".1.2.3";
        elseif fases == "ABCN"
            conex1 = ".1.2.3.0";
        end
    else
        conex1 = conex2;
    end 
    z= string(SSDBT{i,20});
    l= string(SSDBT{i,14});
    fprintf(fid,'New line.%s phases = %s bus1 = BT_%s%s bus2 = BT_%s%s \n~ length = %s units = m \n~ linecode = %s\n\n', x, phase, c, conex1, y, conex2, z, l);
    if i == rows_SSDBT
        clear x c y z i l fases a conex2 conex1 phase PN_2_CONTAIN
    end    
end 
fclose(fid);

fid = fopen(sprintf('%s.txt', N7) ,'wt');
for i = 1:rows_UCMT %Escrita loads MT
    x = string(UCMT{i,1});
    c = string(UCMT{i,2});
    kv = TTEN(TTEN.COD_ID == str2double(UCMT{i,18}), 3);
    kv{1,1} = regexprep(kv{1, 1}, 'V$', '');
    kv{1,1} = regexprep(kv{1, 1}, 'k$', '');
    kv{1,1} = regexprep(kv{1, 1}, ' ', '');
    kv=str2double(kv{1,1});
    soma = sum(UCMT{i,36:47},2) /(30*12*24);
    fases = string(UCMT{i,16});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN"
        conn ="wye";
    elseif fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABC" || fases == "ABCN"
        conn ="delta";
    end
    fprintf(fid,'New load.%s phases = 3 model = 8 ZIPV = [0.5 0 0.5 0 0 1 0.9] daily = A4 bus = %s.1.2.3 kv = %g pf = 0.92 kw = %g  conn = %s\n\n', x, c, kv, soma, conn);
    if i == rows_UCMT
        clear x c y z i kv column soma fases conn
    end 
end   
fclose(fid);

fid = fopen(sprintf('%s.txt', N6) ,'wt');
for i = 1:rows_UCBT %Escrita loads BT
    x = string(UCBT{i,1});
    
    fases = string(UCBT{i,17});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" || fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABN" || fases == "BCN" || fases == "CAN"
        phase ="1";
    elseif fases == "ABC" || fases == "ABCN"
        phase ="3";
    end
    
    c = string(UCBT{i,5});    
    tip_cc = string(UCBT{i,16});
    if contains(tip_cc,"COM-") == 1
        daily = "Comercial";
    elseif contains(tip_cc,"RES-") == 1    
        daily = "Residencial";
    elseif contains(tip_cc,"IND-") == 1    
        daily = "Industrial";
    elseif contains(tip_cc,"SP-") == 1    
        daily = "Serv_Pub";     
    elseif contains(tip_cc,"IP-") == 1    
        daily = "IP";        
    end
    
    PN_2_CONTAIN= find(strcmp(PN_2_UNTRD.Var1,c));    
    if PN_2_CONTAIN ~= '0' 
        if fases == "A"
            conex2 = ".1";
        elseif fases == "B"
            conex2 = ".2";
        elseif fases == "C"
            conex2 = ".3";
        elseif fases == "N"
            conex2 = ".0";            
        elseif fases == "AN"
            conex2 = ".1.0";
        elseif fases == "BN"
            conex2 = ".2.0";
        elseif fases == "CN"
            conex2 = ".3.0";
        elseif fases == "AB"
            conex2 = ".1.2";
        elseif fases == "BC"
            conex2 = ".2.3";
        elseif fases == "CA"
            conex2 = ".3.1";
        elseif fases == "ABN"
            conex2 = ".1.2.0";
        elseif fases == "BCN"
            conex2 = ".2.3.0";
        elseif fases == "CAN"
            conex2 = ".3.1.0";
        elseif fases == "ABC"
            conex2 = ".1.2.3";
        elseif fases == "ABCN"
            conex2 = ".1.2.3.0";
        end
    else
        if fases == "A"
            conex2 = ".1";
        elseif fases == "B"
            conex2 = ".2";
        elseif fases == "C"
            conex2 = ".3";
        elseif fases == "N"
            conex2 = ".4";        
        elseif fases == "AN"
            conex2 = ".1.4";
        elseif fases == "BN"
            conex2 = ".2.4";
        elseif fases == "CN"
            conex2 = ".3.4";
        elseif fases == "AB"
            conex2 = ".1.2";
        elseif fases == "BC"
            conex2 = ".2.3";
        elseif fases == "CA"
            conex2 = ".3.1";
        elseif fases == "ABN"
            conex2 = ".1.2.4";
        elseif fases == "BCN"
            conex2 = ".2.3.4";
        elseif fases == "CAN"
            conex2 = ".3.1.4";
        elseif fases == "ABC"
            conex2 = ".1.2.3";
        elseif fases == "ABCN"
            conex2 = ".1.2.3.4";
        end
    end 
    
    kv = TTEN(TTEN.COD_ID == str2double(UCBT{i,19}), 3);
    kv{1,1} = regexprep(kv{1, 1}, 'V$', '');
    kv{1,1} = regexprep(kv{1, 1}, 'k$', '');
    kv{1,1} = regexprep(kv{1, 1}, ' ', '');
    kv=str2double(kv{1,1});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" 
        kv=kv/1000;
    else
        kv = base_kv_ten_lin;
    end     

    soma = sum(UCBT{i,25:36},2) /(30*12*24);

    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN"
        conn = "wye";
    elseif fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABC" || fases == "ABCN"
        conn = "delta";
    end
    fprintf(fid,'New load.%s phases = %s model = 8 ZIPV = [0.5 0 0.5 0 0 1 0.85] daily = %s bus = BT_%s%s kv = %g pf = 0.92 kw = %g  conn = %s\n\n', x, phase, daily, c, conex2, kv, soma, conn);
    if i == rows_UCBT
        clear x fases phase conex2 soma kv conn tip_cc daily 
    end 
end
fclose(fid);

fid = fopen(sprintf('%s.txt', N9) ,'wt');
for i = 1:rows_PIP %Escrita PIP
    x = string(PIP{i,2});
    
    fases = string(PIP{i,12});
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" || fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABN" || fases == "BCN" || fases == "CAN"
        phase ="1";
    elseif fases == "ABC" || fases == "ABCN"
        phase ="3";
    end

    c = string(PIP{i,10});       
    if fases == "A"
        conex2 = ".1";
    elseif fases == "B"
        conex2 = ".2";
    elseif fases == "C"
        conex2 = ".3";
    elseif fases == "AN"
        conex2 = ".1.4";
    elseif fases == "BN"
        conex2 = ".2.4";
    elseif fases == "CN"
        conex2 = ".3.4";
    elseif fases == "AB"
        conex2 = ".1.2";
    elseif fases == "BC"
        conex2 = ".2.3";
    elseif fases == "CA"
        conex2 = ".3.1";
    elseif fases == "ABN"
        conex2 = ".1.2.4";
    elseif fases == "BCN"
        conex2 = ".2.3.4";
    elseif fases == "CAN"
        conex2 = ".3.1.4";
    elseif fases == "ABC"
        conex2 = ".1.2.3";
    elseif fases == "ABCN"
        conex2 = ".1.2.3.4";
    end
    kv = TTEN(TTEN.COD_ID == str2double(PIP{i,14}), 3);
    kv{1,1} = regexprep(kv{1, 1}, 'V$', '');
    kv{1,1} = regexprep(kv{1, 1}, 'k$', '');
    kv{1,1} = regexprep(kv{1, 1}, ' ', '');
    kv=str2double(kv{1,1});
    
    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN" 
        kv=kv/1000;
    else
        kv = base_kv_ten_lin;
    end     

    soma = sum(PIP{i,20:31},2) /(30*12*24);

    if fases == "A" || fases == "B" || fases == "C" || fases == "AN" || fases == "BN" || fases == "CN"
        conn = "wye";
    elseif fases == "AB" || fases == "BC" || fases == "CA" || fases == "ABN" || fases == "BCN" || fases == "CAN" || fases == "ABC" || fases == "ABCN"
        conn = "delta";
    end
    fprintf(fid,'New load.%s phases = %s model = 8 ZIPV = [0.5 0 0.5 0 0 1 0.85] daily = IP bus = BT_%s%s kv = %g pf = 0.92 kw = %g  conn = %s\n\n', x, phase, c, conex2, kv, soma, conn);
    if i == rows_UCBT
        clear x fases phase conex2 soma kv conn 
    end 
end
fclose(fid);
% //----------------------------------------//