clear all
clc
%% ����Excel�ļ�
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};                
%% ��ʾ������ҵ�x���ļ���Ҫ����ֱ��ǣ�
fprintf('���ҵ�%d���ļ���Ҫ����ֱ��ǣ�',length(fileName));
fprintf('\n');
for i=1:length(fileName)
    fprintf(strcat(fileName{1,i},'\n'));
end
%% ����Ҫ������������ν��д���
for i=1:length(fileName)
    disp(sprintf('���ڴ����%d/%d���ļ�',i,length(fileName)));  %���ڴ����x/X���ļ�
    filename=fileName{1,i};%%�������ݱ�������
    f_dataProcess(filename,length(fileName));
end