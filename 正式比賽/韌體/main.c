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
#define E_INDEX  0x20

#define read_eeprom_byte(address) eeprom_read_byte ((const uint8_t*)address)
#define write_eeprom_byte(address,value) eeprom_write_byte ((uint8_t*)address,(uint8_t)value)

 
static uchar qaz[12]={'0','1','2','3','4','5','6','7','8','9','*','#'};
static uchar picture1[96]={0x1f,0x10,0x10,0x10,0x10,0x10,0x10,0x10,0x1f,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x1f,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x10,0x10,0x10,0x10,0x10,0x10,0x10,0x1f,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x1f,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x1f,0x1f,0x10,0x17,0x17,0x17,0x17,0x17,0x17,0x1f,0x00,0x1f,0x1f,0x1f,0x1f,0x1f,0x1f,0x1f,0x01,0x1d,0x1d,0x1d,0x1d,0x1d,0x1d,0x17,0x17,0x17,0x17,0x17,0x17,0x00,0x1f,0x1f,0x1f,0x1f,0x1f,0x1f,0x1f,0x00,0x1f,0x1d,0x1d,0x1d,0x1d,0x1d,0x1d,0x00,0x1f};
static uchar VB_ON_OFF; //0 : off  1 : on   2 : end
static uchar page,scan_in,count1,station_number,Countdown,TR,xx,qq,w;
static int Hz;
static unsigned short scan_compared,scan_value,btn,password_location;
static uchar hour_math,minute_math,second_math,shift_index,shift_direct;
static uchar scan[4]={0x7D,0x7B,0x77,0x7E};
static uchar password[4], password_in[4];    //password
static uchar temp_timer[6];                  //PC
static uchar temp_timer_setup[6];            //[0][1]時  [2][3]分 [4][5]秒
static uchar temp_timer_rtc[3];              //[0]時  [1]分 [3]秒
static uchar abc[43]={' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ','M','o','v','i','n','g',' ','L','o','g','o',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '}; 
static unsigned long timer[12];
static unsigned long ID;

uchar z1,z2,z3,xy;

void process_keyin(unsigned short value);
void process_keyin_end(unsigned short value);



void LCD_instruction(uchar RS , uchar RW ,uchar b) //RS R/W Data
{	
	RS = RS << 1 ;
	RW = RW << 3 ; 
	PORTD = (PORTD & 0xF5) | RS | RW ;   //要改
	PORTD = PORTD &  ~E_INDEX; 
	_delay_ms(1);		//延遲時間
	PORTD = PORTD |  E_INDEX;          //E  = 1
	_delay_ms(1);
	PORTB = (PORTB & 0xC0) | (b & 0x3F); 
	PORTD = (PORTD & 0x3F) | (b & 0xC0);  
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
		for (a=0 ; a<48 ; a++)
		{
			LCD_CDRAM(0x00 + a) ;
			write_data(picture1[a]) ;
			
			
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

	for(col=0;col<2;col++)   //掃描列 行#de
	{	
		PORTC=scan[col];
		rowkey=PINC & 0x70; //111 0000
		if(rowkey !=0x70)
		{
			switch(rowkey)
			{
			case 0x60:     // 110 0000
			key = key | (0x01<<btndata);
			break;
			case 0x50:   //101 0000
				key = key | (0x02<<btndata);
			break;
			case 0x30: //011 0000
				key = key | (0x04<<btndata);
			break;
			case 0x40:  //100 0000
				key = key | (0x03<<btndata);
			break;
			case 0x20:   // 010  0000
				key = key | (0x05<<btndata);
			break;
			case 0x10:   //001 0000
				key = key | (0x06<<btndata);
			break;
			case 0x00:
				key = key | (0x07<<btndata);
			break;
			}
		}
		btndata=btndata+2;

	}
	return key;
}
uchar VB_btn(void)
{
	uchar vb;
	switch	(Button_scan())
	{
		case 0x01:
			vb=1;
		break;
		case 0x02:
			vb=2;
		break;
		case 0x04:
			vb=3;
		break;
		case 0x08:
			vb=4;
		break;
		case 0x10:
			vb=5;
		break;
		case 0x20:
			vb=6;
		break;
		case 0x40:
			vb=7;
		break;
		case 0x80:
			vb=8;
		break;
		case 0x100:
			vb=9;
		break;
		case 0x200:
			vb=10;
		break;
		case 0x400:
			vb=11;
		break;
		case 0x800:
			vb=12;
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
			}
			

			DISP_Str(0x80,"On Time ");
			DISP_Chr(0x88, 0x30 |  *(data+2));  //時
			DISP_Chr(0x89, 0x30 |  *(data+3));  //
			DISP_Chr(0x8B, 0x30 |  *(data+4));  //分
			DISP_Chr(0x8C, 0x30 |  *(data+5));  //
			DISP_Chr(0x8E, 0x30 |  *(data+6));  //秒
			DISP_Chr(0x8F, 0x30 |  *(data+7));  //
			DISP_Chr(0x8A,':');  
			DISP_Chr(0x8D,':');  
			DISP_Str(0xc0,"Station ");
			DISP_Str(0xcb,"    ");
			DISP_Chr(0xc8,0x30 | *(data) / 100);
			DISP_Chr(0xc9,0x30 | (*(data)%100)  / 10);
			DISP_Chr(0xcA,0x30 | *(data)  % 10);
			
			
		break;
		case 0x05:
			DISP_Chr(0xCf, qaz[*(data+2)]);
		break;
		case 0x06:
			LCD_remove();
			VB_ON_OFF=100;
			page=0;
			hour_math=*(data+2);
			minute_math=*(data+3);
			second_math=*(data+4);
			
		break;
		
		case 0x08:
			page=20;
		break;
		case 0x09:
			page=21;
		break;
		case 0x0A:
			page=22;
		break;
		case 0x0c:
			LCD_remove();
		break;
		case 0x0d:
			
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
	DDRC = 0x0f;    //    0      0      0      0      1      1       1      1
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
	
	/******************* Timer0設定 *********************/
	TIMSK |= (1 << TOIE0);
    TCCR0 |= (1 << CS02);             // set prescaler to 256 and start the timer 
    TCNT0 = 128; 
	


	
	
	Init_LCD();
	//Module();
	
 
 	station_number = 72;    
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
	ID=0;

	
	
    while(1)
	{
		wdt_reset();
		usbPoll();
		
		if (page == 20)
		{
			PORTD=(PORTD & 0xfe) | (qq);
			qq++;
			if (qq==2)
			qq=0;
			
			_delay_ms(2);
			
		}
		else if (page== 21)
		{
			PORTD=(PORTD & 0xfe) | (qq);
			qq++;
			if (qq==2)
			qq=0;
			
			_delay_ms(1);
			
		}
		else if (page== 22)
		{
			PORTD=(PORTD & 0xfe) | (qq);
			qq++;
			if (qq==2)
			qq=0;
			
			_delay_us(500);
			
		}
		if(timer[1] ==0) 
		{
			timer[1]=200;
				if (hour_math<23 && minute_math==59 && second_math==59)
					{
						hour_math++;
						minute_math=0;
						second_math=0;
					}
					else if (hour_math==23 && minute_math==59 && second_math==59)
					{
						hour_math=0;
						minute_math=0;
						second_math=0;
					}
					else if ( minute_math<59 && second_math==59)
					{
						minute_math++;
						second_math=0;
					}
					else if ( minute_math==59 && second_math==59)
					{
						hour_math++;
						minute_math=0;
						second_math=0;
					}
					else if (second_math<59)
						second_math++;
					else if (second_math==59)
					{
						minute_math++;
						second_math=0;
				}
		
		}

		if (VB_ON_OFF != 50)
		{
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
		}
		
	}
	return 0;
}




/********************************************************************************************/
void process_keyin(unsigned short value)
{
	switch(page)
	{
		case 0:
			switch(value)
			{
				case 0x01:
					DISP_Chr(0xcf,'1');
				break;
				case 0x02:
					DISP_Chr(0xcf,'2');
				break;
				case 0x04:
					DISP_Chr(0xcf,'3');
				break;
				case 0x08:
					DISP_Chr(0xcf,'4');
				break;
				case 0x10:
					DISP_Chr(0xcf,'5');
				break;
				case 0x20:
					DISP_Chr(0xcf,'6');
				break;
				case 0x40:
					DISP_Chr(0xcf,'7');
				break;
				case 0x80:
					DISP_Chr(0xcf,'8');
				break;
				case 0x100:
					DISP_Chr(0xcf,'9');
				break;
				case 0x200:
					DISP_Chr(0xcf,'*');
				break;
				case 0x400:
					DISP_Chr(0xcf,'0');
				break;
				case 0x800:
					DISP_Chr(0xcf,'#');
				break;
				case 0xA00:
					LCD_remove();
					page=1;
				break;
			}
			if(timer[2] ==0) 
			{
				timer[2]=90;
				if(ID < 105)
					ID++;
				else
					ID = 1;
				
				
				DISP_Str(0x88,"WELCOME ");
				DISP_Chr(0xca,0x30 |  ID / 100);
				DISP_Chr(0xcb,0x30 | (ID%100) / 10);
				DISP_Chr(0xcc,0x30 | ID % 10);
			}
		break;
		case 1:
			DISP_Str(0x80,"   Time ");
			DISP_Chr(0x88,0x30 | (hour_math /10));
			DISP_Chr(0x89,0x30 | (hour_math %10));
			DISP_Chr(0x8A,':');
			DISP_Chr(0x8B,0x30 | (minute_math /10));
			DISP_Chr(0x8C,0x30 | (minute_math %10));
			DISP_Chr(0x8D,':');
			DISP_Chr(0x8E,0x30 | (second_math /10));
			DISP_Chr(0x8F,0x30 | (second_math %10));
			DISP_Str(0xC0,"Station 072     ");
		break;
	

	}
}

void process_keyin_end(unsigned short value)
{
	
	switch(page)
	{
		case 0:
			DISP_Str(0x80,"   Time ");
			DISP_Chr(0x88,0x30 | (hour_math /10));
			DISP_Chr(0x89,0x30 | (hour_math %10));
			DISP_Chr(0x8A,':');
			DISP_Chr(0x8B,0x30 | (minute_math /10));
			DISP_Chr(0x8C,0x30 | (minute_math %10));
			DISP_Chr(0x8D,':');
			DISP_Chr(0x8E,0x30 | (second_math /10));
			DISP_Chr(0x8F,0x30 | (second_math %10));
			DISP_Str(0xC0,"Station 072     ");
		break;
	}
}
