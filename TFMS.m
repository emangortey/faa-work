 [num,text,raw] = xlsread('TFMS_FLOW_20170421_data');
len = length(raw)
col1 = raw(:,1);
col2 = raw(:,2);
start1=[];
start2=[];
end1=[];
end2=[];
reason1=[];
reason2=[];
ref1=[];
ref2=[];

for i = 1: len
    
    if strfind(col1{i},'Text_Reference')
        ref2 = [ref2 col2(i)];
        end2 = [end2 col2(i-4)];
        start2 = [start2 col2(i-5)];
        
    else
        i =i+1
            
    
        
    end
end
 

    
    

        
 
  

start2 = start2';
end2 =end2';
ref2 = ref2';
vec = [start2 end2 ref2];

% area =[];
% 
% 
% fid = fopen('TFMS_FLOW_20170421_tref.txt','rt')
% %# read the whole file to a temporary cell array
% 
% tmp = textscan(fid,'%s','Delimiter','\n')
% fclose(fid)
% 
% %# remove the lines starting with headerline
% tmp = tmp{1}
% idx = cellfun(@(x) strcmp(x(1:10),'headerline'), tmp)
% tmp(idx) = []
% 
% %# split and concatenate the rest
% result = regexp(tmp,' ','split')
% result = cat(1,result{:})
% 
% %# delete temporary array (if you want)
% clear tmp
% 
% % xlswrite('load.xls',vec);
% area=[];
% reason=[];
% count = 1;
% for j = 1:165
%     if strfind(tmp{j},'CONSTRAINED AREA')
%         long = length(tmp{j})
%         need = tmp{j}
%         area =[area need(long-2:end)
%           R == tmp{j+1}
%           LONG == length(R)
%           reason == [reason R(8:end)]
%           count == count+1 
%     else
%     count = count
%     end
% end
