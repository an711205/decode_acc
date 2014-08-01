clc;
clear all;
close all;

%%���w�즹�@�C
[FileName1 , PathName1] = uigetfile('*.xlsx','');
[data_catch] = xlsread([PathName1 FileName1]);
%%
%{
�`�N�GEXCEL�ɮת��Ĥ@�CŪ���X�ӡA�|��u������X�ɮפ֤@�C�C
29�OŪ�X21�C�L�θ�T�åH�̫e8�C��@�@�է���T�����צ���ѽX���ӱq��B�}�l��(�Ĥ@�խY��n����22~29�C�h����)�C
ans�O�b�̫�~���H4m�C
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