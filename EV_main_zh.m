clear all
clc
%% 载入Excel文件
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};                
%% 提示输出共找到x个文件需要处理分别是：
fprintf('共找到%d个文件需要处理分别是：',length(fileName));
fprintf('\n');
for i=1:length(fileName)
    fprintf(strcat(fileName{1,i},'\n'));
end
%% 将需要处理的数据依次进行处理
for i=1:length(fileName)
    disp(sprintf('正在处理第%d/%d个文件',i,length(fileName)));  %正在处理第x/X个文件
    filename=fileName{1,i};%%输入数据表格的名称
    f_dataProcess(filename,length(fileName));
end