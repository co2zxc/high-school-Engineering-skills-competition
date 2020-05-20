//1.TIMER ���_���[�J,
//2.LCD �ϥܥ��[�J,�ݴ���
//3.�ɶ����վ�

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

/* �����ܼ� */
uchar scan_line, scan_in, scan_value, count1;
uchar page, station_number, shift_index, shift_direct;
uchar temp_timer[6];                  //PC
uchar temp_timer_setup[6];            //[0][1]��  [2][3]�� [4][5]��
uchar temp_timer_rtc[3];              //[0]��  [1]�� [3]��
uchar event, link ; 
unsigned int timer[4];


uchar keyboard_scan(void);
void process_keyin(uchar value);
void process_timer_run(void); 
void process_shift(void);
void process_timer_pc(void);
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
 
    *(data)   = 0;  	// Key value 
                        
 		*(data+1) = 0;		// Reserve (zero)
 		*(data+2) = 0;		// Reserve (zero)
 		*(data+3) = 0;		// Reserve (zero)
 		*(data+4) = 0;		// Reserve (zero)
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
//    		 outputControl(0x20);		    // ---- Generate a Positive pulse
//    		 outputControl(0x0);			  // ----
    	}     	
    	else if(*(data+1) == 0x01)      /* ��s��ܮɶ� */
    	{    	 
         temp_timer[0] = *(data+2);  //��
         temp_timer[1] = *(data+3);  //
         temp_timer[2] = *(data+4);  //��
         temp_timer[3] = *(data+5);  //
         temp_timer[4] = *(data+6);  //��
         temp_timer[5] = *(data+7);  //                                                                                                                                     		 	     		 
    	}
    	else if(*(data+1) == 0x02)      /* ��s��ܮɶ� */
    	{    	 
         temp_timer[0] = *(data+2);  //��
         temp_timer[1] = *(data+3);  //
         temp_timer[2] = *(data+4);  //��
         temp_timer[3] = *(data+5);  //
         temp_timer[4] = *(data+6);  //��
         temp_timer[5] = *(data+7);  //                                                                                                                                                      		 	     		 
    	}    		    
    	else if(*(data+1) == 0x03)      /* �{������ */
    	{    	 
//    	   event = 1;
//         link  = 2;         //�{������                                                                                                                                     		 	     		     	             	                               		 	     		 
    	}	  
    	else if(*(data+1) == 0x04)      /* �]�wSTATION NUMBER */
    	{    	 
    	   station_number = *(data+2);
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
      
   Init_LCD();                        /* LCD��l�]�w */ 
   lcd_set_char(0);                   /* �ۦ�w�q�r�� */
    
   DISP_Str(0x80,"Digital Clock");  
           

   /* �ܼƪ�l�]�w */ 
   scan_line = 0;          //0-2   
   link  = 0;
 	event = 1;  
 	station_number = 37;    
 	count1 = 0; 
   page   = 0;	                               
   
   for(i=0;i<6;i++)
      temp_timer_setup[i] = 0;   
   for(i=0;i<4;i++)
      timer[i] = 0;  

   
   for(;;)          /* main event loop */
   {                
      wdt_reset();
      usbPoll();		
		
		if(timer[0] ==0)                          //5ms x 6 = 30ms, ��L���y�ɶ�
		{                                        
		   timer[0] = 6;    
		   
         scan_value = keyboard_scan();         
         if(scan_value < 12)       
            process_keyin(scan_value);    		   
		}   	
      
      if(timer[1] ==0)                          //5ms x 200 = 1s real time & ��s��ܯ��O
      {
         timer[1] = 200; 
         process_timer_run();                   //�C�����@�� 
         
         process_show_station();                //�C�����@�� 
      }   

      if(timer[2] ==0)                          //5ms x 20 = 100ms ,���ʮɶ�
      {
         timer[2] = 20; 
         process_shift();                       //�C100ms����@��   
      }   

      if(timer[3] ==0)                          //5ms x 40 = 200ms ,PC�ɶ���ܧ�s
      {
         timer[3] = 40; 
         
         process_timer_pc();                    //�C0.5�����@��     
      }
          
                   
   }
   return 0;
}


/* ------------------------------------------------------------------------- */
uchar keyboard_scan(void)                  //��^��=0xff(�L�����J), (0-9 , * = 10, # = 11) 
{
   uchar value;
   scan_in = PINC & 0x07;
   value   = 0xff;
   if(scan_in != 0x07)          //������ʧ@
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
//   PORTD = PORTD | 0x20;        
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

void process_keyin(uchar value)            //�������۹����ʧ@
{
   uchar i;
      
   if(page == 0)   
   {
      if(value == 10)   
      {
         DISP_Str(0x80,"MODE            ");         
         page = 1;         
      }               
   }   
   else if(page == 1)   
   {
      if(value == 1)                           //�ҡB	��1
      {  
      	 for(i=0;i<6;i++)
      	    temp_timer_setup[i] = 0; 
         DISP_Str(0x80,"00:00:00        ");                  
         count1 = 0;                          
         page = 10;         
      }     
      else if(value == 2)                      //�A�B	��2
      {
         DISP_Str(0x80,"  :  :          ");   
         DISP_Chr(0x80, 0x30 | temp_timer_setup[0]);               
         DISP_Chr(0x81, 0x30 | temp_timer_setup[1]);         
         DISP_Chr(0x83, 0x30 | temp_timer_setup[2]);
         DISP_Chr(0x84, 0x30 | temp_timer_setup[3]);
         DISP_Chr(0x86, 0x30 | temp_timer_setup[4]);
         DISP_Chr(0x87, 0x30 | temp_timer_setup[5]);
         page = 20;         
      }       
      else if(value == 3)                      //���B	��3
      {
         DISP_Str(0x80,"number =        ");    //��ܯ��O                  
         DISP_Chr(0x89, 0x30 | (station_number/10));               
         DISP_Chr(0x8A, 0x30 | (station_number%10));         

         page = 30;         
      }       
      else if(value == 4)                      //�B�B	��4
      {
         DISP_Str(0x80,"                ");                            
         shift_index  = 0;
         shift_direct = 0; 
         page = 40;         
      }   
      else if(value == 5)                      //���B	��5
      {
         DISP_Str(0x80,"  :  :          ");   
         DISP_Chr(0x80, 0x30 | temp_timer[0]);               
         DISP_Chr(0x81, 0x30 | temp_timer[1]);         
         DISP_Chr(0x83, 0x30 | temp_timer[2]);
         DISP_Chr(0x84, 0x30 | temp_timer[3]);
         DISP_Chr(0x86, 0x30 | temp_timer[4]);
         DISP_Chr(0x87, 0x30 | temp_timer[5]);
         page = 50;         
      }   
      else if(value == 11)
      {	
         DISP_Str(0x80,"Digital Clock   ");    //������#��^�W�@�h        
         page = 0;            
      }                
   }   
   else if(page == 10)                         //��-�]�w�ɶ�
   {
      if(count1 == 0)                           //��_�Q���
      {  
         if(value < 3)
         {	
            temp_timer_setup[0] = value;
            DISP_Chr(0x80, 0x30 | temp_timer_setup[0]);
            count1++;
         }   
      }  
      else if(count1 == 1)                      //��_�Ӧ�� 
      {  
         if(
         	  ((temp_timer_setup[0] == 2)&&(value < 4))  ||
         	  ((temp_timer_setup[0] <  2)&&(value < 10))
         	 )
         {	
            temp_timer_setup[1] = value;
            DISP_Chr(0x81, 0x30 | temp_timer_setup[1]);             	
            count1++;                        
         }   
      }  
      else if(count1 == 2)                      //��_�Q��� 
      {  
         if(value < 6)
         {	
            temp_timer_setup[2] = value;
            DISP_Chr(0x83, 0x30 | temp_timer_setup[2]);
            count1++;
         }    
      } 
      else if(count1 == 3)                      //��_�Ӧ�� 
      {  
         if(value < 10)
         {	
            temp_timer_setup[3] = value;
            DISP_Chr(0x84, 0x30 | temp_timer_setup[3]);
            count1++;
         }  
      } 
      else if(count1 == 4)                      //��_�Q��� 
      {  
         if(value < 6)
         {	
            temp_timer_setup[4] = value;
            DISP_Chr(0x86, 0x30 | temp_timer_setup[4]);
            count1++;
         }    
      } 
      else if(count1 == 5)                      //��_�Ӧ�� 
      {  
         if(value < 10)
         {	
            temp_timer_setup[5] = value;
            DISP_Chr(0x87, 0x30 | temp_timer_setup[5]);
            count1++;
            temp_timer_rtc[0] = temp_timer_setup[0]*10 + temp_timer_setup[1]; 
            temp_timer_rtc[1] = temp_timer_setup[2]*10 + temp_timer_setup[3]; 
            temp_timer_rtc[2] = temp_timer_setup[4]*10 + temp_timer_setup[4];                         
         }  
      }   
      else if(count1 == 6)                      //������#��^�W�@�h
      {  
         if(value == 11)                     
         {	
            DISP_Str(0x80,"MODE            ");         
            page = 1;            
         }  
      }                                                                                   
   }   
   else if(page == 20)                         //�A-����
   {
      if(value == 11)
      {	
         DISP_Str(0x80,"MODE            ");    //������#��^�W�@�h       
         page = 1;            
      }    	
   }	
   else if(page == 30)                         //��-��ܲէO
   {
      if(value == 11)
      {	
         DISP_Str(0x80,"MODE            ");    //������#��^�W�@�h       
         page = 1;            
      }    	
   }	
   else if(page == 40)                         //�B-�s�i�r�����k����
   {
      if(value == 11)
      {	
         DISP_Str(0x80,"MODE            ");    //������#��^�W�@�h       
         page = 1;            
      }    	
   }   
   else if(page == 50)                         //��-PC����
   {
      if(value == 11)
      {	
         DISP_Str(0x80,"MODE            ");    //������#��^�W�@�h       
         page = 1;            
      }    	
   }      
}

void process_timer_run(void)               //�A-����, page == 20
{	
   if(page == 20)                          //�ɶ����i��
   {
      if(temp_timer_rtc[2] >= 59)          //�B�z��
      {	
         temp_timer_rtc[2] = 0;
         
         if(temp_timer_rtc[1] >= 59)       //�B�z��   
         {
            temp_timer_rtc[1] = 0;   	
            
            if(temp_timer_rtc[0] >= 23)
               temp_timer_rtc[0] = 0;   	               	
            else
               temp_timer_rtc[0]++;		
         }	
         else
            temp_timer_rtc[1]++;	
      }   
      else    	
      	 temp_timer_rtc[2]++;

      DISP_Chr(0x80, 0x30 | (temp_timer_rtc[0]/10));   //��s��ܮɶ�              
      DISP_Chr(0x81, 0x30 | (temp_timer_rtc[0]%10));       	
      DISP_Chr(0x83, 0x30 | (temp_timer_rtc[1]/10));               
      DISP_Chr(0x84, 0x30 | (temp_timer_rtc[1]%10));       	
      DISP_Chr(0x86, 0x30 | (temp_timer_rtc[2]/10));               
      DISP_Chr(0x87, 0x30 | (temp_timer_rtc[2]%10));       	
   }			
}
 
void process_show_station(void)            //��-��ܯ��O
{
   if(page == 30)  
   {   
      DISP_Chr(0x89, 0x30 | (station_number/10));               
      DISP_Chr(0x8A, 0x30 | (station_number%10));    
   }
}

void process_shift(void)                   //�B-�s�i�r�����k����, page == 40
{
   if(page == 40)	
   {
   	DISP_Chr(0x80|shift_index, ' ');   //�M��
      if(shift_direct)	                 // = 0 �k�� ,  = 1 ���� , 
      {	
         if(shift_index == 0)	
            shift_direct = 0;
         else
         	  shift_index--;  
      }   
      else
      {
         if(shift_index >= 15)	
            shift_direct = 1;
         else
         	  shift_index++;  
      }	
      
      if(shift_direct)	
//         DISP_Chr(0x80|shift_index, '<');   	
         DISP_Chr(0x80|shift_index, 0);   	
      else   
//      	DISP_Chr(0x80|shift_index, '>');   	
         DISP_Chr(0x80|shift_index, 0);   	
   }		
}

void process_timer_pc(void)                //��-���PC����, page == 50
{	
   if(page == 50)                          //�ɶ����i��
   {
      DISP_Chr(0x80, 0x30 | temp_timer[0]);     //��s��ܮɶ�           
      DISP_Chr(0x81, 0x30 | temp_timer[1]);         
      DISP_Chr(0x83, 0x30 | temp_timer[2]);
      DISP_Chr(0x84, 0x30 | temp_timer[3]);
      DISP_Chr(0x86, 0x30 | temp_timer[4]);
      DISP_Chr(0x87, 0x30 | temp_timer[5]);      	
   }				
   
}




