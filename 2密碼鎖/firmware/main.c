//1.PC6 掃描
//2.EEPROM 讀寫


/*
This example should run on most AVRs with only little changes. No special
hardware resources except INT0 are used. You may have to change usbconfig.h for
different I/O pins for USB. Please note that USB D+ must be the INT0 pin, or
at least be connected to INT0 as well.
*/
#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

#include "lcd.h" 
#include "lcd.c" 

#define read_eeprom_byte(address) eeprom_read_byte ((const uint8_t*)address)
#define write_eeprom_byte(address,value) eeprom_write_byte ((uint8_t*)address,(uint8_t)value)

/* 全域變數 */
uchar scan_line, scan_in, scan_value, count1;
uchar page, station_number, flag_delay;
uchar event, link ; 
unsigned int timer[4];
uchar password[4], password_in[4];    //password

uchar keyboard_scan(void);
void process_keyin(uchar value);
void process_show_station(void);


/* ----------------------------- USB interface ----------------------------- */
PROGMEM const char usbHidReportDescriptor[22] = {    /* USB report descriptor */
    0x06, 0x00, 0xff,              // USAGE_PAGE (Generic Desktop)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    //   LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              //   LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    //   REPORT_SIZE (8)
    0x95, 0x80,                    //   REPORT_COUNT (128)
    0x09, 0x00,                    //   USAGE (Undefined)
    0xb2, 0x02, 0x01,              //   FEATURE (Data,Var,Abs,Buf)
    0xc0                           // END_COLLECTION
};
/* Since we define only one feature report, we don't use report-IDs (which
 * would be the first byte of the report). The entire report consists of 128
 * opaque data bytes.
 */

/* The following variables store the status of the current data transfer */
static uchar    currentAddress;
static uchar    bytesRemaining;

/* usbFunctionRead() is called when the host requests a chunk of data from
 * the device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionRead(uchar *data, uchar len)
{
   if(len > bytesRemaining)
      len = bytesRemaining;
 	 if(currentAddress == 0)
 	 {
 
      *(data)   = 0x30 | read_eeprom_byte(0);		                            
      *(data+1) = 0x30 | read_eeprom_byte(1);		
 		*(data+2) = 0x30 | read_eeprom_byte(2);				
 		*(data+3) = 0x30 | read_eeprom_byte(3);			
 		
 		*(data+4) = 0;				 		 		
 		*(data+5) = 0;		// Reserve (zero)
 		*(data+6) = 0;		// Reserve (zero)
 		*(data+7) = 0;		// Reserve (zero)
   }
   currentAddress += len;
   bytesRemaining -= len;
   
   return len;
}

/* usbFunctionWrite() is called when the host sends a chunk of data to the
 * device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionWrite(uchar *data, uchar len)
{
// 	 static uchar tempB,tempD;
	 
   if(bytesRemaining == 0)
      return 1;               /* end of transfer */
   if(len > bytesRemaining)
      len = bytesRemaining;
      
   if(currentAddress == 0)
   {   	
    	if((*(data+1) & 0x20) == 0x20)
    	{
    	   
    	}     			     
    	else if(*(data+1) == 0x04)      /* 設定STATION NUMBER */
    	{    	 
    	   station_number = *(data+2);
         flag_delay = 1; 
    	}	      	
    	else if(*(data+1) == 0x05)      /* 更改密碼 */
    	{    	     	   
    	   password_in[0] = *(data+2);
    	   password_in[1] = *(data+3);
    	   password_in[2] = *(data+4);
    	   password_in[3] = *(data+5);    	       	       	   
         write_eeprom_byte(0,*(data+2));     //save     	     	       	       	       	   
         write_eeprom_byte(1,*(data+3));     //save     	     	       	       	       	   
         write_eeprom_byte(2,*(data+4));     //save     	     	       	       	       	   
         write_eeprom_byte(3,*(data+5));     //save     	     	       	       	       	                              
    	}	
 	}
   currentAddress += len;
   bytesRemaining -= len;
   return bytesRemaining == 0; /* return 1 if this was the last chunk */
}

usbMsgLen_t usbFunctionSetup(uchar data[8])
{
	usbRequest_t    *rq = (void *)data;
    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionRead() to obtain data */
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* since we have only one report type, we can ignore the report-ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* use usbFunctionWrite() to receive data from host */
        }
    }else{
        /* ignore vendor type requests, we don't use any */
    }
    return 0;
}
/* ------------------------------------------------------------------------- */
ISR (TIMER0_OVF_vect)                       // timer0 overflow interrupt
{
   uchar   i;
   // event to be exicuted every 5ms here
   TCNT0 = 22;  
      
   for(i=0;i<4;i++)
   {
      if(timer[i]>0)
         timer[i]--;           
   }    
}

int main(void)
{	
	 uchar   i;	
   wdt_enable(WDTO_1S);
   /* Even if you don't use the watchdog, turn it off here. On newer devices,
     * the status of the watchdog (on/off, period) is PRESERVED OVER RESET!
   */
     
   /* IO SETUP */  
   {
     	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
     	DDRB = 0x3f;    //    0      0      1      1      1      1      1      1
     	                //    x      x     out    out    out    out    out    out
                      //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
     	PORTB = 0x05;   //    0      0      0      0      0      0      0      0
  
     
     				          //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
     	DDRC = 0x78;    //    0      1      1      1      1      0      0      0  
     					        //  (Read)  out    out    out    out    in      in    in
                      //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
      PORTC = 0x07;   //    0      0      0      0      0      0      0      0
  
     	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
     	DDRD = 0xeb;    //    1      1      1      0      1      0      1      1 
     					        //   out    out    out    D-     out     D+    out    out
     	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
     	PORTD = 0x0;    //    0      0      0      0      0      0      0      0  
   }
	
   usbInit();
   usbDeviceDisconnect();  /* enforce re-enumeration, do this while interrupts are disabled! */
	 i = 0;
   while(--i)              /* fake USB disconnect for > 250 ms */
   {             
      wdt_reset();
      _delay_ms(1);
   }
   usbDeviceConnect();
   sei();
   
   TIMSK |= (1 << TOIE0);
   TCCR0 |= (1 << CS02);             // set prescaler to 256 and start the timer 
   TCNT0 = 128; 
      
   Init_LCD();                        /* LCD初始設定 */ 
//   lcd_set_char(0);                   /* 自行定義字型 */
               
   /* 變數初始設定 */ 
   scan_line = 0;          //0-2   
   link  = 0;
 	event = 1;  
 	station_number = 37;    
 	count1 = 0; 
   page   = 0;	 
   flag_delay = 0;   

   DISP_Str(0x80,"Input Code:     ");   
   DISP_Str(0xC0,"Station =");             //顯示站別                  
   DISP_Chr(0xC9, 0x30 | (station_number/10));               
   DISP_Chr(0xCA, 0x30 | (station_number%10));
   
   for(i=0;i<4;i++)
      timer[i] = 0;
   for(i=0;i<4;i++)    
   {  
      password[i] = 0;          //開機預設為0  
      password_in[i] = 0;
   }   
   
   for(;;)          /* main event loop */
   {                
      wdt_reset();
      usbPoll();		
		
  		if(timer[0] ==0)                          //5ms x 6 = 30ms, 鍵盤掃描時間
  		{                                        
  		   timer[0] = 6;    
  		   
           scan_value = keyboard_scan();         
           if(scan_value < 12)       
              process_keyin(scan_value);    		   
  		}   	
        
//      if(timer[1] ==0)                          //5ms x 200 = 更新顯示站別
//      {
//         timer[1] = 200; 
//         
//         process_show_station();                //每秒執行一次 
//      }       

      if((timer[2] ==0)&&(flag_delay == 1))     //
      {
         flag_delay = 0; 
         DISP_Str(0x80,"Input Code:     ");
//         DISP_Str(0xC0,"                ");        
    
         DISP_Str(0xC0,"Station =");             //顯示站別                  
         DISP_Chr(0xC9, 0x30 | (station_number/10));               
         DISP_Chr(0xCA, 0x30 | (station_number%10));
      }                     
   }
   return 0;
}


/* ------------------------------------------------------------------------- */
uchar keyboard_scan(void)                  //返回值=0xff(無按鍵輸入), (0-9 , * = 10, # = 11) 
{
   uchar value;
   scan_in = PINC & 0x07;
   value   = 0xff;
   if(scan_in != 0x07)          //有按鍵動作
   {
      scan_in = (~scan_in)& 0x07;  
      if (scan_in == 0x01)
         value = 1;
      else if (scan_in == 0x02)
         value = 2;
      else if (scan_in == 0x04)
         value = 3;
      
      value = value + (scan_line * 3);   
      if(value == 11)
         value = 0;
      else if(value == 12)      
         value = 11;
   }   
   else
      value = 0xff;   
   
   PORTC = PORTC | 0x78;     
   PORTD = PORTD | 0x20;        
   if(scan_line == 3)
   {
      scan_line = 0;
      PORTC = PORTC & 0xf7;            
   }   
   else if(scan_line == 0)
   {
      scan_line = 1;
      PORTC = PORTC & 0xEF;            
   }   
   else if(scan_line == 1)
   {
      scan_line = 2;
      PORTC = PORTC & 0xDF;            
   } 
   else if(scan_line == 2)
   {
      scan_line = 3;            
//      PORTD = PORTD & 0xDF;            
      PORTC = PORTC & 0xBF;            
   }    
   
   return value;
}

void process_keyin(uchar value)            //執行按鍵相對應動作
{
   uchar i;
      
   if(page == 0)   
   {
         
      if((value < 10)&&(count1<4))         //輸入密碼範圍0-9                
      {  
         DISP_Chr((0x8B + count1), (0x30 | value));                          
         password_in[count1] = value;             //紀錄輸入密碼
         count1++;
         if(count1==4)                            //輸入完畢  
         { 
              for(i=0;i<4;i++) 
                 password[i] = read_eeprom_byte(i); //read             
                 
         	  if((password_in[0] == password[0]) && //執行比對
         	     (password_in[1] == password[1]) &&
         	     (password_in[2] == password[2]) &&
         	     (password_in[3] == password[3]))           	
         	  {
         	     DISP_Str(0xC0,"Good Job!       ");	         	     
         	  } 
         	  else
         	  {
         	     DISP_Str(0xC0,"Sorry !!        ");	                      	              	  	                                                   
         	  }	  

            for(i=0;i<4;i++)         
               password_in[i] = 0;
         	  count1 = 0;
         	  
            //_delay_ms(2000);                      //延遲顯示  
            flag_delay = 1;  
            timer[2] = 400;                      	           	              
         }	
      }   	
      else if(value == 10)   
      {
         DISP_Str(0x80,"Save Code :     ");      
         DISP_Str(0xC0,"                ");              
         count1 = 0;
         page   = 1;         
      }  
      else if(value == 20)                 //按『*』+「1」顯示組別 
      {
         page   = 30; 
         DISP_Str(0x80,"Number :        ");           
         DISP_Chr(0x89, 0x30 | (station_number/10));               
         DISP_Chr(0x8A, 0x30 | (station_number%10));   
         
         count1 = 0;            	
      }	      	      
   }   
   else if(page == 1)   
   {
      if(value < 10)                           //輸入密碼範圍0-9                
      {  
      	 if(count1<4)
      	 {	
            password[count1] = value;          //儲存密碼      	   
            DISP_Chr((0x8B + count1), (0x30 | value));                          
            count1++;
          }             
      }               
      else if(value == 10)                     //組別
      {
         DISP_Str(0x80,"number =        ");    //顯示站別 
         DISP_Chr(0x89, 0x30 | (station_number/10));               
         DISP_Chr(0x8A, 0x30 | (station_number%10));           
         page = 30;         
      }      
      else if((value == 11)&&(count1==4))      //#儲存
      {	
         DISP_Str(0xC0,"Save Ok!        ");    //完畢用#返回上一層                          
         
         for(i=0;i<4;i++) 
            write_eeprom_byte(i,password[i]);     //save 
                                                   
         for(i=0;i<4;i++)         
            password_in[i] = 0;
         count1 = 0;                              
         page = 0;  
         flag_delay = 1;  
         timer[2] = 400;        
      }                
   }   
   else if(page == 30)                     //顯示組別
   {
      if(value == 10)
      {	
         DISP_Str(0x80,"Input Code:     ");    //完畢用#返回上一層                         
         page = 0;            
      }    	
   }	
    
}

 
void process_show_station(void)            //顯示站別
{
   if(page == 30)  
   {   
      DISP_Chr(0x89, 0x30 | (station_number/10));               
      DISP_Chr(0x8A, 0x30 | (station_number%10));    
   }
}


