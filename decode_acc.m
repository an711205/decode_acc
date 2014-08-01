clc;
clear all;
close all;

%%明安到此一遊
[FileName1 , PathName1] = uigetfile('*.xlsx','');
[data_catch] = xlsread([PathName1 FileName1]);
%%
%{
注意：EXCEL檔案的第一列讀不出來，會比真正的輸出檔案少一列。
29是讀出21列無用資訊並以最前8列當作一組完整訊號長度估算解碼應該從何處開始解(第一組若剛好為第22~29列則不算)。
ans是在最後才乘以4m。
%}

for i =1:29
    
    if(data_catch(i,2) == 255)
        
        count_255 = i;
        
    end
    
end

for j = 1:floor((length(data_catch)-count_255)/8)
    
    %x
    if(data_catch(count_255+1+8*(j-1),2) <= 127)
        
        ans(1,1+j-1) = ( data_catch(count_255+1+8*(j-1),2)*16 + data_catch(count_255+1+1+8*(j-1),2)/16 );
        
    elseif(data_catch(count_255+1+8*(j-1),2) > 127)
        
        ans(1,1+j-1) = (-1)*( (255-data_catch(count_255+1+8*(j-1),2))*16 + (255-data_catch(count_255+1+1+8*(j-1),2)+1)/16 );
        
    end
    
    %y
    if(data_catch(count_255+1+2+8*(j-1),2) <= 127)
        
        ans(2,1+j-1) = ( data_catch(count_255+1+2+8*(j-1),2)*16 + data_catch(count_255+1+3+8*(j-1),2)/16 );
        
    elseif(data_catch(count_255+1+2+8*(j-1),2) > 127)
        
        ans(2,1+j-1) = (-1)*( (255-data_catch(count_255+1+2+8*(j-1),2))*16 + (255-data_catch(count_255+1+3+8*(j-1),2)+1)/16 );
        
    end
    
    %z
    if(data_catch(count_255+1+4+8*(j-1),2) <= 127)
        
        ans(3,1+j-1) = ( data_catch(count_255+1+4+8*(j-1),2)*16 + data_catch(count_255+1+5+8*(j-1),2)/16 );
        
    elseif(data_catch(count_255+1+4+8*(j-1),2) > 127)
        
        ans(3,1+j-1) = (-1)*( (255-data_catch(count_255+1+4+8*(j-1),2))*16 + (255-data_catch(count_255+1+5+8*(j-1),2)+1)/16 );
        
    end
    
end

ans = ans .* 4 * 10^(-3);