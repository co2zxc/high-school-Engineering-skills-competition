#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */
#define RS_INDEX 0x01 
#define RW_INDEX 0x02
#define E_INDEX  0x08

#define read_eeprom_byte(address) eeprom_read_byte ((const uint8_t*)address)
#define write_eeprom_byte(address,value) eeprom_write_byte ((uint8_t*)address,(uint8_t)value)

 

static uchar picture1[8]={0x11,0x11,0x0A,0x04,0x1F,0x1F,0x1F,0x00};
static uchar VB_ON_OFF; //0 : off  1 : on   2 : end
static uchar page,scan_in,count1,station_number,Countdown,Gb,TR,xx;

static unsigned short scan_compared,scan_value,btn,password_location;
static uchar hour_math,minute_math,second_math,shift_index,shift_direct;
static uchar scan[4]={0x5f,0x6f,0x77,0x3f};
static uchar moving_loge[1];
static uchar password[4], password_in[4];    //password
static uchar temp_timer[6];                  //PC
static uchar temp_timer_setup[6];            //[0][1]時  [2][3]分 [4][5]秒
static uchar temp_timer_rtc[3];              //[0]時  [1]分 [3]秒
static uchar abc[43]={' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','M','o','v','i','n','g',' ','L','o','g','o',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '}; 
static unsigned long timer[12];
void process_keyin(unsigned short value);
void process_keyin_end(unsigned short value);
void process_timer_pc(unsigned short value);




void LCD_instruction(uchar RS , uchar RW ,uchar b) //RS R/W Data
{	
	
	RW = RW << 1 ; 
	PORTD = (PORTD & 0xFC) | RS | RW ;
	PORTD = PORTD &  ~E_INDEX; 
	_delay_ms(1);		//延遲時間
	PORTD = PORTD |  E_INDEX;          //E  = 1
	_delay_ms(1);
	PORTB = (PORTB & 0xC0) | (b & 0x3F); 
	PORTD = (PORTD & 0x3F) | (b & 0xC0); 
	_delay_ms(1);	
	PORTD = (PORTD & 0xF7); 
	_delay_ms(1);						// 一個短時間的延遲時序	
	PORTD = PORTD & ~E_INDEX;                 //E  = 0	
	_delay_us(1); 
}

void LCD_remove()   //清除顯示
 {
	LCD_instruction(0,0,0x01); 
}

void LCD_show(uchar D, uchar C, uchar B) // D C B 顯示控制
{
	D = D << 2 ;
	C = C << 1 ;
	LCD_instruction(0,0,0x08 | D | C | B); 
}

void Init_LCD()           //LCD 前置作業 ps:一定要打
{
	_delay_ms(15);       //延遲時間
	LCD_instruction(0,0,0x30); //功能設定
	_delay_ms(5);        //延遲時間
	LCD_instruction(0,0,0x3C); //功能設定(8位元)
	LCD_instruction(0,0,0x08); //另顯示器OFF
	LCD_remove();     //清除顯示    
	LCD_instruction(0,0,0x06); //輸入模式設定
	LCD_show(1,0,0);   //另顯示器ON
}

void LCD_Move(uchar SC, uchar LR) // S/C R/L 游標移位 或 整個顯示移位
{
	SC = SC << 3 ;
	LR = LR << 2 ;
	LCD_instruction(0,0,0x10 | SC | LR);  
}





//==========================設定DDRAM的位址 功能:設定下一個要存入的資料DDRAM之位址 ==============================================================
void LCD_DDRAM(uchar rD)  
{
	LCD_instruction(0,0,0x80 | rD);
}



//==========================設定CDRAM的位址  功能:設定下一個要存入的資料CGRAM之位址 =============================================================
void LCD_CDRAM(uchar sC) 
{
	LCD_instruction(0,0,0x40 | sC); 
}

	
//==========================打資料寫入DDRAM 或 CGRAM     功能: 1.把字元碼寫入DDRAM內,以便顯示出相對應之字元 2. 把自創之圖形存入CGRAM內 =========
void write_data(uchar wDC)
{
	LCD_instruction(1,0,0x00| wDC); 
}


//==========================讀取DDRAM 或 CGRAM之內容 ==========================================================================================    
void read_contents(uchar rDC)     
{
    LCD_instruction(1,1,0x00 | rDC); 
}


//========= 把圖形碼存入CGRAM內 ======================
void Module()
{
	uchar a;
		for (a=0 ; a<8 ; a++)
		{
			LCD_CDRAM(0x00 + a) ;
			write_data(picture1[a]) ;
			_delay_ms(1);
			
		}
}


//==================  顯示字元在指定 DDRAM 的位子 ======
void DISP_Chr(uchar DDRAM,uchar Data) // DDRAM 位置      Data 輸入字元資料
{
	LCD_DDRAM(DDRAM);
	write_data(Data);
} 


// Clock
void outputControl()
{
	PORTD = PORTD | 0x20; //0x20 00100000
	PORTD = PORTD & 0xDF; //0xDF 11011111
}


//=================== 按鈕掃描 =========================================
unsigned short Button_scan(void)
{
	uchar col = 0,rowkey= 0,btndata=0;
	unsigned short key=0;

	for(col=0;col<4;col++)
	{	
		PORTC=scan[col];
		rowkey=PINC & 0x07;
		if(rowkey !=0x07)
		{
			switch(rowkey)
			{
			case 0x03:
			key = key | (0x01<<btndata);
				break;
			case 0x05:
				key = key | (0x02<<btndata);
			break;
			case 0x06:
				key = key | (0x04<<btndata);
			break;
			case 0x01:
				key = key | (0x03<<btndata);
			break;
			case 0x02:
				key = key | (0x05<<btndata);
			break;
			case 0x04:
				key = key | (0x06<<btndata);
			break;
			case 0x00:
				key = key | (0x07<<btndata);
			break;
			}
		}
		btndata=btndata+3;

	}
	return key;
}
uchar VB_btn(void)
{
	uchar vb;
	switch	(Button_scan())
	{
		case 0x200:
			vb=1;
		break;
		case 0x800:
			vb=2;
		break;
		case 0xA00:
			vb=3;
		break;
	}
	return vb;
}
void DISP_Str(char addr1,char *str)
{
		LCD_DDRAM(addr1);         // 設定DD RAM位址第一行第一列
		while(*str !=0)
	    write_data(*str++);	// 呼叫顯示字串函式
}




/* ------------------------------------------------------------------------- */
/* ----------------------------- USB interface ----------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {    /* USB report descriptor */
    0x06, 0x00, 0xff,              // USAGE_PAGE (Generic Desktop)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    // LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              // LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    // REPORT_SIZE (8)
    0x95, 0x80,                    // REPORT_COUNT (128)
    0x09, 0x00,                    // USAGE (Undefined)
    0xb2, 0x02, 0x01,              // FEATURE (Data,Var,Abs,Buf)
    0xc0                           // END_COLLECTION
};

/* Since we define only one feature report, we don't use report-IDs (which
 * would be the first byte of the report). The entire report consists of 128
 * opaque data bytes.
 */

/* The following variables store the status of the current data transfer */
uchar    currentAddress;
uchar    bytesRemaining;

/* ------------------------------------------------------------------------- */

/* usbFunctionRead() is called when the host requests a chunk of data from
 * the device. For more information see the documentation in usbdrv/usbdrv.h.
 */
uchar   usbFunctionRead(uchar *data, uchar len)
{
//ReadData
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = VB_btn();	// Key value
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
	if(bytesRemaining == 0)
        return 1;               /* end of transfer */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){

	// OutDataEightByte
	
    switch(*(data+1))
	{
		
		case 0x04:
			if (VB_ON_OFF!=50)
			{
				LCD_remove();
				VB_ON_OFF=50;
				page=50;
			}
			
			DISP_Str(0x80," Current Time:  ");
			DISP_Str(0xc0, "      "); 
			DISP_Str(0xcb, "     ");
			DISP_Chr(0xc6, 0x30 |  *(data+2));  //時
			DISP_Chr(0xc7, 0x30 |  *(data+3));  //
			DISP_Chr(0xc9, 0x30 |  *(data+4));  //分
			DISP_Chr(0xcA, 0x30 |  *(data+5));  //
			
			if (TR == 0)
			{
			DISP_Chr(0xc8,':');  
			TR++;
			}
			else
			{
			DISP_Chr(0xc8,' '); 
			TR=0;
			}
			
			
		break;
		
		case 0x05:
			DISP_Str(0x80," Current Second: ");
			DISP_Str(0xc6, "  ");  //秒
			DISP_Str(0xcb, "     ");
			DISP_Chr(0xc9, 0x30 |  *(data+6));  //秒
			DISP_Chr(0xcA, 0x30 |  *(data+7));  //
			
			if (TR == 0)
			{
			DISP_Chr(0xc8,':');  
			TR++;
			}
			else
			{
			DISP_Chr(0xc8,' '); 
			TR=0;
			}
		break;	
		
		case 0x06:
			LCD_remove();
			VB_ON_OFF=100;
			page=0;

			
		break;
		
		case 0x07:
			station_number = *(data+2);
			DISP_Str(0x80," Station Number:");
			if (station_number>=10)
				{
					DISP_Str(0xC0,"        ");
					DISP_Str(0xCA,"      ");
					DISP_Chr(0xc8,0x30 | station_number / 10);
					DISP_Chr(0xc9,0x30 | station_number % 10);
				}
				else if (station_number<10)
				{
					DISP_Str(0xC0,"         ");
					DISP_Str(0xCA,"      ");
					DISP_Chr(0xc9,0x30 | station_number);

				}
		break;	
			
		case 0x08:
			DISP_Chr(0xc0,abc[xx]);
			DISP_Chr(0xc1,abc[xx+1]);
			DISP_Chr(0xc2,abc[xx+2]);
			DISP_Chr(0xc3,abc[xx+3]);
			DISP_Chr(0xc4,abc[xx+4]);
			DISP_Chr(0xc5,abc[xx+5]);
			DISP_Chr(0xc6,abc[xx+6]);
			DISP_Chr(0xc7,abc[xx+7]);
			DISP_Chr(0xc8,abc[xx+8]);
			DISP_Chr(0xc9,abc[xx+9]);
			DISP_Chr(0xca,abc[xx+10]);
			DISP_Chr(0xcb,abc[xx+11]);
			DISP_Chr(0xcc,abc[xx+12]);
			DISP_Chr(0xcd,abc[xx+13]);
			DISP_Chr(0xce,abc[xx+14]);
			DISP_Chr(0xcf,abc[xx+15]);
			xx++;
			
			if(xx==27)
			xx=0;
		break;	

		case 0x25:			
		break;
			
		case 0x26:			
		break;
		
		
		case 0x27:			
		break;
		
		
		default:
		break;		
	}
	
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* return 1 if this was the last chunk */
	}
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

ISR (TIMER0_OVF_vect)                       // timer0 overflow interrupt
{
   uchar   i;
   // event to be exicuted every 5ms here
   TCNT0 = 22;  
      
   for(i=0;i<12;i++)
   {
      if(timer[i]>0)
         timer[i]--;           
   }    
}




int main(void)
{

	uchar i;
    wdt_enable(WDTO_1S);
					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3F;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
	PORTB = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x78;    //    0      1      1      1      1      0       0     0
					//  (Read)  reset   out    out    out    out    out    out
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0
    PORTC = 0x7f;   //    0      1      1      1      1      1      1      1
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xEB;    //    1      1      1      0      1      0      1      1
					//   out    out    out    D-     out    D+     out    out
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------
					

	
    usbInit();
    usbDeviceDisconnect();  /* enforce re-enumeration, do this while interrupts are disabled! */
	i = 0;
    

	while(--i){             /* fake USB disconnect for > 250 ms */
        wdt_reset();
        _delay_ms(0.1);
    }
    usbDeviceConnect();
    sei();
	_delay_ms(250);
	
	TIMSK |= (1 << TOIE0);
    TCCR0 |= (1 << CS02);             // set prescaler to 256 and start the timer 
    TCNT0 = 128; 
	Init_LCD();
	Module();
	
 
 	station_number = 6;    
 	count1 = 0; 
    page   = 0;	        
	for(i=0;i<12;i++)
	timer[i] = 0;
	
	password[0]=1;password[1]=1;password[2]=0;password[3]=2;
	
	for(i=0;i<4;i++)
	password_in[i]=0;
	
	VB_ON_OFF=0;
	hour_math=0;								
	minute_math=0;
	second_math=0;
	password_location=0;
	
	/* main event loop */
    while(1)
	{
		wdt_reset();
		usbPoll();
		
		if(timer[0] ==0)                          //5ms x 6 = 30ms, 鍵盤掃描時間
		{                                        
			timer[0] = 6; 
		    btn=Button_scan();
			if(btn!=scan_compared)
				scan_compared=btn;
			else
				btn=0x1000;
			
			if (VB_ON_OFF==0)
			process_keyin(btn);
			
			if (VB_ON_OFF==100)
			process_keyin_end(Button_scan());
			
			
		}
			
		
		/* if(timer[3] ==0)                          //5ms x 40 = 200ms ,PC時間顯示更新
		{
			timer[3] = 40; 
			process_timer_pc(btn);                    //每0.5秒執行一次     
		}
*/
	}
	return 0;
}




/********************************************************************************************/
void process_keyin(unsigned short value)
{
	switch(page)
	{
		case 0:
			DISP_Chr(0x80,0x00);
			DISP_Str(0x81,"Input Password:");
			page=20;
		break;
		
		case 20:
			
			if (password_location==4)
				password_location=0;
			
			switch(value)
			{

				case 0x001:
					DISP_Chr((0xC0+password_location),'1');
					password_in[password_location]=1;
					password_location++;
				break;
				case 0x002:
					DISP_Chr((0xC0+password_location),'2');
					password_in[password_location]=2;
					password_location++;
				break;
				case 0x004:
					DISP_Chr((0xC0+password_location),'3');
					password_in[password_location]=3;
					password_location++;
				break;
				case 0x008:
					DISP_Chr((0xC0+password_location),'4');
					password_in[password_location]=4;
					password_location++;
				break;
				case 0x010:
					DISP_Chr((0xC0+password_location),'5');
					password_in[password_location]=5;
					password_location++;
				break;
				case 0x20:
					DISP_Chr((0xC0+password_location),'6');
					password_in[password_location]=6;
					password_location++;
				break;
				case 0x40:
					DISP_Chr((0xC0+password_location),'7');
					password_in[password_location]=7;
					password_location++;
				break;
				case 0x80:
					DISP_Chr((0xC0+password_location),'8');
					password_in[password_location]=8;
					password_location++;
				break;
				case 0x100:
					DISP_Chr((0xC0+password_location),'9');
					password_in[password_location]=9;
					password_location++;
				break;
				case 0x400:
					DISP_Chr((0xC0+password_location),'0');
					password_in[password_location]=0;
					password_location++;
				break;
			}
			if(password_location==4)
			{
				
					if((password_in[0] == password[0]) && //執行比對
					(password_in[1] == password[1]) &&
					(password_in[2] == password[2]) &&
					(password_in[3] == password[3])) 
					{

						for(Gb=0;Gb<4;Gb++)
						password_in[Gb]=0;						
						DISP_Str(0xc0,"Password OK!    ");
						page=99;
						
					}
					else
					{
						
						DISP_Str(0xc0,"Password Error! ");
						_delay_ms(500);
						LCD_remove();
						page=0;
						
						for(Gb=0;Gb<4;Gb++)
						password_in[Gb]=0;
					
					}
					password_location=0;
			}
		break;
		
		case 21:
			switch(value)
			{
				case 0x01:
					LCD_remove();
					DISP_Str(0x80,"Station Number:");
					if (station_number>=100)
					{
						
						DISP_Chr(0xc0,0x30 | station_number / 100);
						DISP_Chr(0xc1,0x30 | ((station_number-100) / 10));
						DISP_Chr(0xc2,0x30 | ((station_number-100) % 10));
					}
					else if (station_number>10)
					{
						DISP_Chr(0xc0,0x30 | station_number / 10);
						DISP_Chr(0xc1,0x30 | station_number % 10);
					}
					else if (station_number<10)
					{
						DISP_Chr(0xc0,0x30 | station_number);

					}
					page=22;
				break;
				case 0x02:
					LCD_remove();
					DISP_Chr(0x80,0x00);
					timer[11]=100;
					page=23;
					
				break;
				case 0x03:
					page=0;
				break;
			}
		break;
		case 22:
			if	(value==0)
			{
				LCD_remove();
				DISP_Str(0x80,"1-Number 2-Moving");
				DISP_Str(0xc0,"Loge 3-Exit");
				page=21;
					
			}
			
		break;
		case 23:
			
			if (value==0)
			{
				LCD_remove();
				DISP_Str(0x80,"1-Number 2-Moving");
				DISP_Str(0xc0,"Loge 3-Exit");
				shift_index=0;
				page=21;
						
			}
			else if (value==0x1000)
			{
				if (timer[11]==0)
				{
					timer[11]=100;
					if(shift_index<15)
					{
						LCD_Move(1,1);
						shift_index++;
					}
					else 
					{
						page=24;
						shift_index=0;
					}
				}
			}
			
		break;
		
		case 24:
			if (value==0)
			{
				LCD_remove();
				DISP_Str(0x80,"1-Number 2-Moving");
				DISP_Str(0xc0,"Loge 3-Exit");
				shift_direct=0;
				page=21;
						
			}
			else if (value==0x1000)
			{
				if (timer[11]==0)
				{
					timer[11]=100;
					if(shift_direct<15)
					{
						LCD_Move(1,0);
						shift_direct++;
					}
					else 
					{
						page=23;
						shift_direct=0;
					}
				}
			}
			
		break;
		
		
	}
}
void process_timer_pc(unsigned short value)                // page == 50
{	
   if(page == 50)                          //時間的進行
   {
		DISP_Chr(0x86, 0x30 | temp_timer[0]);     //更新顯示時間           
		DISP_Chr(0x87, 0x30 | temp_timer[1]);         
		DISP_Chr(0x88, ':'); 
		DISP_Chr(0x89, 0x30 | temp_timer[2]);
		DISP_Chr(0x8A, 0x30 | temp_timer[3]);
		DISP_Chr(0x8b, ':'); 
		DISP_Chr(0x8C, 0x30 | temp_timer[4]);
		DISP_Chr(0x8D, 0x30 | temp_timer[5]);  
		
	
	}
}
void process_keyin_end(unsigned short value)
{
	_delay_ms(100);
	switch(page)
	{
		case 0:
			switch(value)
			{
				case 0:
					DISP_Str(0x80," Current Time:  ");
					DISP_Str(0xc6,"00:00");
				break;
				case 0x200:
					DISP_Str(0x80," Current Second:");
					DISP_Str(0xc6,"  :00");					
				break;
				case 0x800:
					DISP_Str(0x80," Station Number:");
					DISP_Str(0xc6,"   00");					
				break;
			}
		break;
	}
}
