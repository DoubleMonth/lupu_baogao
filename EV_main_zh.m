clear all
clc
%% word����
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%�������ݱ�������
i = find('.'==filename);
imname = filename(1:i-1); %% imnameΪ������׺�ļ����� 

string = strcat(imname,'���ݷ�������'); %%��ɴ�excle�ļ�����podu�ļ���
% string ='���ݷ�������';
doc_f='.doc';
spwd=[pwd '\'];
file_name =[spwd string doc_f];
%file_name = [string doc_f]
try
    Word=actxGetRunningServer('Word.Application');
catch
    Word = actxserver('Word.Application');
    
end;

set(Word, 'Visible', 1);
documents = Word.Documents;
if exist(file_name,'file')
    document = invoke(documents,'Open',file_name);
else
    document = invoke(documents, 'Add');
 %   document = invoke(document,'SaveAs');
%     document.SaveAs(file_name);
end
 %document = invoke(documents,'Open',file_name);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
content = document.Content;
duplicate = content.Duplicate;
inlineshapes = content.InlineShapes;
selection= Word.Selection;
paragraphformat = selection.ParagraphFormat;
shape=document.Shapes;

%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%ҳ������
document.PageSetup.TopMargin = 60;
document.PageSetup.BottomMargin = 45;
document.PageSetup.LeftMargin = 45;
document.PageSetup.RightMargin = 45;
set(content, 'Start',0);
set(content, 'Text',string);
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
% selection.Font.Size=50;
[title_a,title_b]=size(string);
rr=document.Range(0,title_b);%ѡ���ı�
rr.Font.Size=20;%�����ı�����
%rr.Font.Bold=4;%�����ı�����

end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.MoveDown;
selection.TypeParagraph;

%% ���ݼ���
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%�������ݱ�������
[excelData,str] = xlsread(fileName{1,1},1);               %��ȡԭʼ���ݱ��е����ݣ�strΪ���ݱ��е��ַ���dataΪ���ݱ��е�����
[excelRow,excelColumn] = size(excelData);        %%��ȡ���ݱ��е����и���
value =  zeros(excelRow,4);                      %����һ����Ӧ������1�еľ������ڴ洢����������
runTime = excelData(excelRow,1); %�ɼ�ʱ��
runDistance = (excelData(excelRow,7)-excelData(1,7))/1000; % ��ʻ����
maxVelocity = max(excelData(:,6));                  %��߳���
averageVelocity = mean(excelData(:,6)); % ƽ������
maxPitch = tand(max(excelData(:,2)))*100;     %����¶�
averagePitch = tand(mean(excelData(:,2)))*100;    %ƽ���¶�

%% ����ռ��
[sortVelocity,index] = sort(excelData(:,6));
speed_num =  zeros(9,1);  % ����ռ�ȸ���
for i=1:excelRow
    if sortVelocity(i,1)<10
        speed_num(1,1) = speed_num(1,1)+1;
    elseif sortVelocity(i,1)<20
        speed_num(2,1) = speed_num(2,1)+1;
    elseif sortVelocity(i,1)<30
        speed_num(3,1) = speed_num(3,1)+1;
    elseif sortVelocity(i,1)<40
        speed_num(4,1) = speed_num(4,1)+1;
    elseif sortVelocity(i,1)<50
        speed_num(5,1) = speed_num(5,1)+1;
    elseif sortVelocity(i,1)<60
        speed_num(6,1) = speed_num(6,1)+1;
    elseif sortVelocity(i,1)<70
        speed_num(7,1) = speed_num(7,1)+1;
    elseif sortVelocity(i,1)<80
        speed_num(8,1) = speed_num(8,1)+1;
    elseif sortVelocity(i,1)<90
        speed_num(9,1) = speed_num(9,1)+1;
    end
end
%% ���¶�ռ��

[sortPitch,index] = sort(excelData(:,2));
sortPitch2(:,1)= tand(sortPitch(:,1))*100;
pitch_num =  zeros(6,1);  % ����ռ�ȸ���
for i=1:excelRow
    if sortPitch2(i,1)<4
        pitch_num(1,1) = pitch_num(1,1)+1;
    elseif sortPitch2(i,1)<8
        pitch_num(2,1) = pitch_num(2,1)+1;
    elseif sortPitch2(i,1)<12
        pitch_num(3,1) = pitch_num(3,1)+1;
    elseif sortPitch2(i,1)<16
        pitch_num(4,1) = pitch_num(4,1)+1;
    elseif sortPitch2(i,1)<20
        pitch_num(5,1) = pitch_num(5,1)+1;
    else
        pitch_num(6,1) = pitch_num(6,1)+1;
    end
end
%% ��� ˵��
selection.MoveDown;
selection.TypeParagraph;
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
set(selection, 'Text','1. �����ռ���������������ͳ�Ʊ��������ʾ��');
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;

selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
% selection.TypeParagraph;
set(selection, 'Text','��1�� �ɼ����ݷ���');
selection.Font.Size=8;
selection.MoveDown;
Tables=document.Tables.Add(selection.Range,22,3);
DTI=document.Tables.Item(1);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';
column_width=[80.575,70.7736,60.7736];
row_height=[20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849,20.5849];
for i=1:3
DTI.Columns.Item(i).Width=column_width(i);
end
for i=1:22
DTI.Rows.Item(i).Height =row_height(i);
end
for i=1:22
for j=1:3
      DTI.Cell(i,j).VerticalAlignment='wdCellAlignVerticalCenter';
end
end
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.TypeParagraph;
DTI.Cell(1,1).Range.Text = '���';%����Ҫ����
DTI.Cell(2,1).Range.Text = '�ɼ�ʱ��(s)';
DTI.Cell(3,1).Range.Text = '��ʻ����(km)';
DTI.Cell(4,1).Range.Text = '��߳���(km/h)';
DTI.Cell(5,1).Range.Text = 'ƽ������(km/h)';
DTI.Cell(6,1).Range.Text = '����ռ��';
DTI.Cell(6,2).Range.Text = '0-10(km/h)';
DTI.Cell(7,2).Range.Text = '10-20(km/h)';
DTI.Cell(8,2).Range.Text = '20-30(km/h)';
DTI.Cell(9,2).Range.Text = '30-40(km/h)';
DTI.Cell(10,2).Range.Text = '40-50(km/h)';
DTI.Cell(11,2).Range.Text = '50-60(km/h)';
DTI.Cell(12,2).Range.Text = '60-70(km/h)';
DTI.Cell(13,2).Range.Text = '70-80(km/h)';
DTI.Cell(14,2).Range.Text = '80-90(km/h)';
DTI.Cell(15,1).Range.Text = '����¶�(deg)';
DTI.Cell(16,1).Range.Text = 'ƽ���¶�(deg)';
DTI.Cell(17,1).Range.Text = '�¶�ռ��';
DTI.Cell(17,2).Range.Text = '0-4(deg)';
DTI.Cell(18,2).Range.Text = '4-8(deg)';
DTI.Cell(19,2).Range.Text = '8-12(deg)';
DTI.Cell(20,2).Range.Text = '12-16(deg)';
DTI.Cell(21,2).Range.Text = '16-20(deg)';
DTI.Cell(22,2).Range.Text = '����20(deg)';
DTI.Cell(1,3).Range.Text = '��ֵ';%����Ҫ����
DTI.Cell(2,3).Range.Text = num2str(runTime);%�ɼ�ʱ��
DTI.Cell(3,3).Range.Text = num2str(runDistance);%��ʻ����
DTI.Cell(4,3).Range.Text = num2str(maxVelocity);%��߳���
DTI.Cell(5,3).Range.Text = num2str(averageVelocity);%ƽ������
DTI.Cell(15,3).Range.Text = num2str(maxPitch);%����¶�
DTI.Cell(16,3).Range.Text = num2str(averagePitch);%ƽ���¶�
DTI.Cell(6,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(1,1)/excelRow)*100));%����ռ��  
DTI.Cell(7,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(2,1)/excelRow)*100));%����ռ��
DTI.Cell(8,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(3,1)/excelRow)*100));%����ռ��
DTI.Cell(9,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(4,1)/excelRow)*100));%����ռ��
DTI.Cell(10,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(5,1)/excelRow)*100));%����ռ��
DTI.Cell(11,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(6,1)/excelRow)*100));%����ռ��
DTI.Cell(12,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(7,1)/excelRow)*100));%����ռ��
DTI.Cell(13,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(8,1)/excelRow)*100));%����ռ��
DTI.Cell(14,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(9,1)/excelRow)*100));%����ռ��
DTI.Cell(17,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(1,1)/excelRow)*100));%�¶�ռ�� 
DTI.Cell(18,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(2,1)/excelRow)*100));%�¶�ռ��
DTI.Cell(19,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(3,1)/excelRow)*100));%�¶�ռ��
DTI.Cell(20,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(4,1)/excelRow)*100));%�¶�ռ��
DTI.Cell(21,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(5,1)/excelRow)*100));%�¶�ռ��
DTI.Cell(22,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(6,1)/excelRow)*100));%�¶�ռ��
%% �ϲ���Ԫ��
DTI.Cell(1, 1).Merge(DTI.Cell(1, 2));
DTI.Cell(2, 1).Merge(DTI.Cell(2, 2));
DTI.Cell(3, 1).Merge(DTI.Cell(3, 2));
DTI.Cell(4, 1).Merge(DTI.Cell(4, 2));
DTI.Cell(5, 1).Merge(DTI.Cell(5, 2));
DTI.Cell(6, 1).Merge(DTI.Cell(14, 1));
DTI.Cell(15, 1).Merge(DTI.Cell(15, 2));
DTI.Cell(16, 1).Merge(DTI.Cell(16, 2));
DTI.Cell(17, 1).Merge(DTI.Cell(22, 1));
%% ͼƬ ˵��
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
selection.TypeParagraph;
set(selection, 'Text','2. �����ռ�����������������������ʾ��');
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;
%% ����ʱ������
plot(excelData(:,1),excelData(:,6),'r-');
grid on;%%��ʾ������
legend('����ʱ������');
title('����ʱ������');
xlabel('ʱ��');
ylabel('����');
%% �������ɵ�����ͼ  
pngFile = strcat(imname,'����ʱ��.png'); %%��ɴ�excle�ļ�����podu�ļ���
figFile = strcat(imname,'����ʱ��.fig'); %%��ɴ�excle�ļ�����podu�ļ���
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% ��ͼ��ճ������ǰ�ĵ���
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% ͼ1 ˵��
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','ͼ1�� ����ʱ������');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% �¶�ʱ������
plot(excelData(:,1),excelData(:,2),'r-');
grid on;%%��ʾ������
legend('�¶�ʱ������');
title('�¶�ʱ������');
xlabel('ʱ��');
ylabel('�¶�');
%% �������ɵ�����ͼ  
pngFile = strcat(imname,'�¶�ʱ��.png'); %%��ɴ�excle�ļ�����podu�ļ���
figFile = strcat(imname,'�¶�ʱ��.fig'); %%��ɴ�excle�ļ�����podu�ļ���
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% ��ͼ��ճ������ǰ�ĵ���
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% ͼ2 ˵��
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','ͼ2�� �¶�ʱ������');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% ������ʻ��������
distanceData = zeros(excelRow,1);
distanceData(:,1) = excelData(:,7)/1000;
plot(distanceData(:,1),excelData(:,6),'r-');
grid on;%%��ʾ������
legend('������ʻ��������');
title('������ʻ��������');
xlabel('��ʻ����');
ylabel('����');
%% �������ɵ�����ͼ  
pngFile = strcat(imname,'������ʻ����.png'); %%��ɴ�excle�ļ�����podu�ļ���
figFile = strcat(imname,'������ʻ����.fig'); %%��ɴ�excle�ļ�����podu�ļ���
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% ��ͼ��ճ������ǰ�ĵ���
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% ͼ3 ˵��
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','ͼ3�� ������ʻ��������');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% �¶���ʻ��������
Pitch2(:,1)= tand(excelData(:,2))*100;
h = plot(distanceData(:,1),Pitch2(:,1),'r-');
grid on;%%��ʾ������
legend('�¶���ʻ��������');
title('�¶���ʻ��������');
xlabel('��ʻ����');
ylabel('�¶�');
%% �������ɵ�����ͼ  
pngFile = strcat(imname,'�¶���ʻ����.png'); %%��ɴ�excle�ļ�����podu�ļ���
figFile = strcat(imname,'�¶���ʻ����.fig'); %%��ɴ�excle�ļ�����podu�ļ���
saveas(gcf,pngFile);
saveas(gcf,figFile);
% hgexport(h, '-clipboard');
print -dbitmap
%��ͼ��ճ������ǰ�ĵ���
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.MoveDown;
%% ͼ4 ˵��
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','ͼ4�� �¶���ʻ����');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
close;

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
document = invoke(document,'SaveAs',file_name); % �����ĵ�
Word.Quit; % �ر��ĵ�