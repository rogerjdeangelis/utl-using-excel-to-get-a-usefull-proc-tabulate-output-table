# utl-using-excel-to-get-a-usefull-proc-tabulate-output-table
Using excel to get a usefull proc tabulate output table  
    Using excel to get a useful proc tabulate output table                                                                                                
                                                                                                                                                          
       Method                                                                                                                                             
           a. send proc tabulate output to excel;                                                                                                         
                                                                                                                                                          
           b. Define macro variable by looking at excel output                                                                                            
                                                                                                                                                          
              %let column_count         = 14;                                                                                                             
              %let right_margin_columns = 2;                                                                                                              
              %let header_range         = xel.'tab$A1:N4'n;   /* headings  */                                                                             
              %let table_range          = xel.'tab$A5:N999'n; /* body */                                                                                  
                                                                                                                                                          
           c. input exel header then body                                                                                                                 
                                                                                                                                                          
                                                                                                                                                          
    This is a manual process, since tabulate does not honor ODS table output.                                                                             
    You hould be able to adjust the code or other tabulate output.                                                                                        
    Protection against being painted into a corner?                                                                                                       
                                                                                                                                                          
    github                                                                                                                                                
    https://tinyurl.com/w2h27py                                                                                                                           
    https://github.com/rogerjdeangelis/utl-using-excel-to-get-a-usefull-proc-tabulate-output-table                                                        
                                                                                                                                                          
    SAS Forum                                                                                                                                             
    https://tinyurl.com/tu8bjf6                                                                                                                           
    https://communities.sas.com/t5/SAS-Procedures/Proc-tabulate-selecting-one-level-of-a-variable-to-be-printed/m-p/608522                                
                                                                                                                                                          
    *_                   _                                                                                                                                
    (_)_ __  _ __  _   _| |_                                                                                                                              
    | | '_ \| '_ \| | | | __|                                                                                                                             
    | | | | | |_) | |_| | |_                                                                                                                              
    |_|_| |_| .__/ \__,_|\__|                                                                                                                             
            |_|                                                                                                                                           
    ;                                                                                                                                                     
                                                                                                                                                          
    * look at the excel output in process below and set these macro variables;                                                                            
                                                                                                                                                          
    %let column_count         = 14;                                                                                                                       
    %let right_margin_columns = 2;                                                                                                                        
    %let header_range         = xel.'tab$A1:N4'n;   /* headings  */                                                                                       
    %let table_range          = xel.'tab$A5:N999'n; /* body */                                                                                            
                                                                                                                                                          
    data have;                                                                                                                                            
       set sashelp.cars(keep = make type origin drivetrain                                                                                                
                where=(type in ('Sedan','Truck') and                                                                                                      
                  make in ('Toyota','Dodge','Ford')));                                                                                                    
    run;quit;                                                                                                                                             
                                                                                                                                                          
    WORK.HAVE total obs=44                                                                                                                                
                                                                                                                                                          
       MAKE     TYPE     ORIGIN    DRIVETRAIN                                                                                                             
                                                                                                                                                          
      Dodge     Sedan     USA        Front                                                                                                                
      Dodge     Sedan     USA        Front                                                                                                                
      Ford      Sedan     USA        Front                                                                                                                
      Ford      Sedan     USA        Rear                                                                                                                 
      Toyota    Sedan     Asia       Rear                                                                                                                 
      Toyota    Sedan     Asia       Front                                                                                                                
      Toyota    Sedan     Asia       Front                                                                                                                
    ....                                                                                                                                                  
                                                                                                                                                          
    *            _               _                                                                                                                        
      ___  _   _| |_ _ __  _   _| |_                                                                                                                      
     / _ \| | | | __| '_ \| | | | __|                                                                                                                     
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                                      
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                                     
                    |_|                                                                                                                                   
    ;                                                                                                                                                     
                                                                                                                                                          
    WORK.WANT total obs=6                                                                                                                                 
                                                                                                                                                          
                          FRONT_     FRONT_     FRONT_     FRONT_     REAR_      REAR_     REAR_     REAR_                                                
     TYPE      MAKE       ASIA_N    ASIA_PCT     USA_N    USA_PCT    ASIA_N    ASIA_PCT    USA_N    USA_PCT                                               
                                                                                                                                                          
     Sedan    Dodge          0           0         7         47         0           0        0          0                                                 
              Ford           0           0         8         53         0           0        3        100                                                 
              Toyota        16         100         0          0         0           0        0          0                                                 
     Truck    Dodge          0           0         0          0         0           0        3         60                                                 
              Ford           0           0         0          0         0           0        2         40                                                 
              Toyota         0           0         0          0         2         100        0          0                                                 
                                                                                                                                                          
                                                                                                                                                          
     *                                                                                                                                                    
     _ __  _ __ ___   ___ ___  ___ ___                                                                                                                    
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                                                   
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                                                                   
    | .__/|_|  \___/ \___\___||___/___/                                                                                                                   
    |_|                                                                                                                                                   
               _        _                         _                                                                                                       
      __ _    | |_ __ _| |__     _____  _____ ___| |                                                                                                      
     / _` |   | __/ _` | '_ \   / _ \ \/ / __/ _ \ |                                                                                                      
    | (_| |_  | || (_| | |_) | |  __/>  < (_|  __/ |                                                                                                      
     \__,_(_)  \__\__,_|_.__/   \___/_/\_\___\___|_|                                                                                                      
                                                                                                                                                          
    ;                                                                                                                                                     
                                                                                                                                                          
    ods excel file="d:\xls\tabx.xlsx" style=minimal options(sheet_name='tab');                                                                            
      proc tabulate data=sashelp.cars(where=(type in ('Sedan','Truck') and make in ('Toyota','Dodge','Ford')));                                           
        format _numeric_ 6.;                                                                                                                              
        class make type origin drivetrain;                                                                                                                
        table (type='Type')*(make='Make'), (drivetrain=' ')*(origin=' ')*(n*f=3. pctn<make>*f=6.)/rts=16 nocellmerge printmiss misstext='0' box="  ";     
        keylabel PctN="Pct" N="N";                                                                                                                        
    run;quit;                                                                                                                                             
    ods excel close;                                                                                                                                      
                                                                                                                                                          
    Up to 40 obs WORK.TST total obs=4                                                                                                                     
                                                                                                                                                          
    Obs     F1      F2     F3      F4    F5     F6    F7       F8    F9     F10    F11     F12    F13    F14                                              
                                                                                                                                                          
     1                     All                        Front                        Rear                                                                   
     2                     Asia          USA          Asia           USA           Asia           USA                                                     
     3                     n       %     n      %     n        %     n       %     n        %     n       %                                               
     4     Type    Make                                                                                                                                   
                                                                                                                                                          
    *    _       __                                                                                                                                       
      __| | ___ / _|  _ __ ___   __ _  ___ _ __ ___  ___                                                                                                  
     / _` |/ _ \ |_  | '_ ` _ \ / _` |/ __| '__/ _ \/ __|                                                                                                 
    | (_| |  __/  _| | | | | | | (_| | (__| | | (_) \__ \                                                                                                 
     \__,_|\___|_|   |_| |_| |_|\__,_|\___|_|  \___/|___/                                                                                                 
                                                                                                                                                          
    ;                                                                                                                                                     
                                                                                                                                                          
    * just in case;                                                                                                                                       
    %symdel                                                                                                                                               
       column_count                                                                                                                                       
       right_margin_columns                                                                                                                               
       header_range                                                                                                                                       
       table_range                                                                                                                                        
    ;                                                                                                                                                     
                                                                                                                                                          
    proc datasets nolist;                                                                                                                                 
      delete want;                                                                                                                                        
    run;quit;                                                                                                                                             
    *          _                                                                                                                                          
      ___ ___ | |  _ __   __ _ _ __ ___   ___  ___   _ __  _   _ _ __ ___  ___                                                                            
     / __/ _ \| | | '_ \ / _` | '_ ` _ \ / _ \/ __| | '_ \| | | | '_ ` _ \/ __|                                                                           
    | (_| (_) | | | | | | (_| | | | | | |  __/\__ \ | | | | |_| | | | | | \__ \                                                                           
     \___\___/|_| |_| |_|\__,_|_| |_| |_|\___||___/ |_| |_|\__,_|_| |_| |_|___/                                                                           
                                                                                                                                                          
    ;                                                                                                                                                     
    libname xel "d:\xls\tabx.xlsx" header=no;                                                                                                             
                                                                                                                                                          
    * look at the excel output and set these macro variables;                                                                                             
                                                                                                                                                          
    %let column_count         = 14;                                                                                                                       
    %let right_margin_columns = 2;                                                                                                                        
    %let header_range         = xel.'tab$A1:N4'n;                                                                                                         
    %let table_range          = xel.'tab$A5:N999'n;                                                                                                       
                                                                                                                                                          
    data want;                                                                                                                                            
                                                                                                                                                          
      if _n_=0 then do; %let rc=%sysfunc(dosubl('                                                                                                         
         /* get colmn names */                                                                                                                            
         data _null_;                                                                                                                                     
                                                                                                                                                          
            array fs[*] $32 f1-f&column_count;                                                                                                            
                                                                                                                                                          
            set &header_range end=dne;                                                                                                                    
            retain fx1-fx&column_count;                                                                                                                   
            length hdr $1000;                                                                                                                             
            array fx[*] $32 fx:;                                                                                                                          
                                                                                                                                                          
            do idx=1 to &column_count;                                                                                                                    
              if fs[idx]="" and idx> (&right_margin_columns + 1) then fs[idx] = fs[idx-1];                                                                
              if fs[idx] not in :(" ","A0"x)       then fx[idx] = catx('_',fx[idx],fs[idx]);                                                              
              hdr=catx(" ",hdr,fx[idx]);                                                                                                                  
            end;                                                                                                                                          
            call symputx("hdr",hdr);                                                                                                                      
                                                                                                                                                          
          run;quit;                                                                                                                                       
          '));                                                                                                                                            
          %array(newnames,values=&hdr);                                                                                                                   
          %array(oldnames,values=1-&column_count);                                                                                                        
          /*                                                                                                                                              
          %put &=hdr;                                                                                                                                     
                                                                                                                                                          
             HDR=Type                                                                                                                                     
                 Make                                                                                                                                     
                 All_Asia_n                                                                                                                               
                 All_Asia_Pct                                                                                                                             
                 All_USA_n                                                                                                                                
                 All_USA_Pct                                                                                                                              
                 Front_Asia_n                                                                                                                             
                 Front_Asia_Pct                                                                                                                           
                 Front_USA_n                                                                                                                              
                 Front_USA_Pct                                                                                                                            
                 Rear_Asia_n                                                                                                                              
                 Rear_Asia_Pct                                                                                                                            
                 Rear_USA_n                                                                                                                               
                 Rear_USA_Pct                                                                                                                             
          */                                                                                                                                              
      end;                                                                                                                                                
                                                                                                                                                          
      set xel.'tab$A5:N999'n(rename=(%do_over(oldnames newnames,phrase=%str(f?oldnames = ?newnames))));                                                   
    run;quit;                                                                                                                                             
                                                                                                                                                          
    libname xel clear;                                                                                                                                    
                                                                                                                                                          
                                                                                                                                                          
    NOTE: There were 6 observations read from the data set XEL.'tab$A5:N999'n.                                                                            
    NOTE: The data set WORK.WANT has 6 observations and 14 variables.                                                                                     
    NOTE: DATA statement used (Total process time):                                                                                                       
          real time           3.48 seconds                                                                                                                
                                                                                                                                                          
