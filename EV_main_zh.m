clear all
clc
%% word设置
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%输入数据表格的名称
i = find('.'==filename);
imname = filename(1:i-1); %% imname为不带后缀文件名称 

string = strcat(imname,'数据分析报告'); %%组成带excle文件名的podu文件名
% string ='数据分析报告';
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
%页面设置
document.PageSetup.TopMargin = 60;
document.PageSetup.BottomMargin = 45;
document.PageSetup.LeftMargin = 45;
document.PageSetup.RightMargin = 45;
set(content, 'Start',0);
set(content, 'Text',string);
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
% selection.Font.Size=50;
[title_a,title_b]=size(string);
rr=document.Range(0,title_b);%选择文本
rr.Font.Size=20;%设置文本字体
%rr.Font.Bold=4;%设置文本字体

end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.MoveDown;
selection.TypeParagraph;

%% 数据计算
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%输入数据表格的名称
[excelData,str] = xlsread(fileName{1,1},1);               %读取原始数据表中的数据：str为数据表中的字符，data为数据表中的数据
[excelRow,excelColumn] = size(excelData);        %%获取数据表中的行列个数
value =  zeros(excelRow,4);                      %建立一个相应行数，1列的矩阵用于存储计算后的数据
runTime = excelData(excelRow,1); %采集时间
runDistance = (excelData(excelRow,7)-excelData(1,7))/1000; % 行驶距离
maxVelocity = max(excelData(:,6));                  %最高车速
averageVelocity = mean(excelData(:,6)); % 平均车速
maxPitch = tand(max(excelData(:,2)))*100;     %最大坡度
averagePitch = tand(mean(excelData(:,2)))*100;    %平均坡度

%% 求车速占比
[sortVelocity,index] = sort(excelData(:,6));
speed_num =  zeros(9,1);  % 车速占比个数
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
%% 求坡度占比

[sortPitch,index] = sort(excelData(:,2));
sortPitch2(:,1)= tand(sortPitch(:,1))*100;
pitch_num =  zeros(6,1);  % 车速占比个数
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
%% 表格 说明
selection.MoveDown;
selection.TypeParagraph;
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
set(selection, 'Text','1. 根据收集的数据做出数据统计表格如下所示：');
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;

selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
% selection.TypeParagraph;
set(selection, 'Text','表1： 采集数据分析');
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
DTI.Cell(1,1).Range.Text = '项次';%不需要更改
DTI.Cell(2,1).Range.Text = '采集时间(s)';
DTI.Cell(3,1).Range.Text = '行驶距离(km)';
DTI.Cell(4,1).Range.Text = '最高车速(km/h)';
DTI.Cell(5,1).Range.Text = '平均车速(km/h)';
DTI.Cell(6,1).Range.Text = '车速占比';
DTI.Cell(6,2).Range.Text = '0-10(km/h)';
DTI.Cell(7,2).Range.Text = '10-20(km/h)';
DTI.Cell(8,2).Range.Text = '20-30(km/h)';
DTI.Cell(9,2).Range.Text = '30-40(km/h)';
DTI.Cell(10,2).Range.Text = '40-50(km/h)';
DTI.Cell(11,2).Range.Text = '50-60(km/h)';
DTI.Cell(12,2).Range.Text = '60-70(km/h)';
DTI.Cell(13,2).Range.Text = '70-80(km/h)';
DTI.Cell(14,2).Range.Text = '80-90(km/h)';
DTI.Cell(15,1).Range.Text = '最大坡度(deg)';
DTI.Cell(16,1).Range.Text = '平均坡度(deg)';
DTI.Cell(17,1).Range.Text = '坡度占比';
DTI.Cell(17,2).Range.Text = '0-4(deg)';
DTI.Cell(18,2).Range.Text = '4-8(deg)';
DTI.Cell(19,2).Range.Text = '8-12(deg)';
DTI.Cell(20,2).Range.Text = '12-16(deg)';
DTI.Cell(21,2).Range.Text = '16-20(deg)';
DTI.Cell(22,2).Range.Text = '大于20(deg)';
DTI.Cell(1,3).Range.Text = '数值';%不需要更改
DTI.Cell(2,3).Range.Text = num2str(runTime);%采集时间
DTI.Cell(3,3).Range.Text = num2str(runDistance);%行驶距离
DTI.Cell(4,3).Range.Text = num2str(maxVelocity);%最高车速
DTI.Cell(5,3).Range.Text = num2str(averageVelocity);%平均车速
DTI.Cell(15,3).Range.Text = num2str(maxPitch);%最大坡度
DTI.Cell(16,3).Range.Text = num2str(averagePitch);%平均坡度
DTI.Cell(6,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(1,1)/excelRow)*100));%车速占比  
DTI.Cell(7,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(2,1)/excelRow)*100));%车速占比
DTI.Cell(8,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(3,1)/excelRow)*100));%车速占比
DTI.Cell(9,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(4,1)/excelRow)*100));%车速占比
DTI.Cell(10,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(5,1)/excelRow)*100));%车速占比
DTI.Cell(11,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(6,1)/excelRow)*100));%车速占比
DTI.Cell(12,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(7,1)/excelRow)*100));%车速占比
DTI.Cell(13,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(8,1)/excelRow)*100));%车速占比
DTI.Cell(14,3).Range.Text = num2str(sprintf('%2.2f%%', (speed_num(9,1)/excelRow)*100));%车速占比
DTI.Cell(17,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(1,1)/excelRow)*100));%坡度占比 
DTI.Cell(18,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(2,1)/excelRow)*100));%坡度占比
DTI.Cell(19,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(3,1)/excelRow)*100));%坡度占比
DTI.Cell(20,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(4,1)/excelRow)*100));%坡度占比
DTI.Cell(21,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(5,1)/excelRow)*100));%坡度占比
DTI.Cell(22,3).Range.Text = num2str(sprintf('%2.2f%%', (pitch_num(6,1)/excelRow)*100));%坡度占比
%% 合并单元格
DTI.Cell(1, 1).Merge(DTI.Cell(1, 2));
DTI.Cell(2, 1).Merge(DTI.Cell(2, 2));
DTI.Cell(3, 1).Merge(DTI.Cell(3, 2));
DTI.Cell(4, 1).Merge(DTI.Cell(4, 2));
DTI.Cell(5, 1).Merge(DTI.Cell(5, 2));
DTI.Cell(6, 1).Merge(DTI.Cell(14, 1));
DTI.Cell(15, 1).Merge(DTI.Cell(15, 2));
DTI.Cell(16, 1).Merge(DTI.Cell(16, 2));
DTI.Cell(17, 1).Merge(DTI.Cell(22, 1));
%% 图片 说明
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
selection.TypeParagraph;
set(selection, 'Text','2. 根据收集的数据现做出曲线如下所示：');
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;
%% 车速时间曲线
plot(excelData(:,1),excelData(:,6),'r-');
grid on;%%显示网格线
legend('车速时间曲线');
title('车速时间曲线');
xlabel('时间');
ylabel('车速');
%% 保存生成的折线图  
pngFile = strcat(imname,'车速时间.png'); %%组成带excle文件名的podu文件名
figFile = strcat(imname,'车速时间.fig'); %%组成带excle文件名的podu文件名
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% 将图形粘贴到当前文档里
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% 图1 说明
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','图1： 车速时间曲线');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% 坡度时间曲线
plot(excelData(:,1),excelData(:,2),'r-');
grid on;%%显示网格线
legend('坡度时间曲线');
title('坡度时间曲线');
xlabel('时间');
ylabel('坡度');
%% 保存生成的折线图  
pngFile = strcat(imname,'坡度时间.png'); %%组成带excle文件名的podu文件名
figFile = strcat(imname,'坡度时间.fig'); %%组成带excle文件名的podu文件名
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% 将图形粘贴到当前文档里
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% 图2 说明
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','图2： 坡度时间曲线');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% 车速行驶距离曲线
distanceData = zeros(excelRow,1);
distanceData(:,1) = excelData(:,7)/1000;
plot(distanceData(:,1),excelData(:,6),'r-');
grid on;%%显示网格线
legend('车速行驶距离曲线');
title('车速行驶距离曲线');
xlabel('行驶距离');
ylabel('车速');
%% 保存生成的折线图  
pngFile = strcat(imname,'车速行驶距离.png'); %%组成带excle文件名的podu文件名
figFile = strcat(imname,'车速行驶距离.fig'); %%组成带excle文件名的podu文件名
saveas(gcf,pngFile);
saveas(gcf,figFile);
%% 将图形粘贴到当前文档里
print -dbitmap
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
%% 图3 说明
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','图3： 车速行驶距离曲线');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
%% 坡度行驶距离曲线
Pitch2(:,1)= tand(excelData(:,2))*100;
h = plot(distanceData(:,1),Pitch2(:,1),'r-');
grid on;%%显示网格线
legend('坡度行驶距离曲线');
title('坡度行驶距离曲线');
xlabel('行驶距离');
ylabel('坡度');
%% 保存生成的折线图  
pngFile = strcat(imname,'坡度行驶距离.png'); %%组成带excle文件名的podu文件名
figFile = strcat(imname,'坡度行驶距离.fig'); %%组成带excle文件名的podu文件名
saveas(gcf,pngFile);
saveas(gcf,figFile);
% hgexport(h, '-clipboard');
print -dbitmap
%将图形粘贴到当前文档里
selection.Range.Paste;
end_of_doc = get(content,'end');
set(selection,'Start',end_of_doc);
selection.MoveDown;
%% 图4 说明
selection.MoveDown;
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
selection.TypeParagraph;
set(selection, 'Text','图4： 坡度行驶距离');
selection.Font.Size=8;
selection.MoveDown;
selection.TypeParagraph;
close;

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
document = invoke(document,'SaveAs',file_name); % 保存文档
Word.Quit; % 关闭文档