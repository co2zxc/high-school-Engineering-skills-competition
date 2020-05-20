#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

#include "Time_Number.h"
#include "Math_Number.h"
#include "Moving_Logo.h"
 

static uchar LCD_airplane[23];
static uchar scan[4]={0xfd,0xfb,0xf7,0xfe};
static unsigned short btn ;




void LCD_instruction(uchar RS , uchar RW ,uchar b) //RS R/W Data
{	 
	RW = RW << 1 ; 
	PORTD = (PORTD & 0xFC) | RW | RS ; 
	_delay_ms(1);                         // 延遲時間
	PORTD = (PORTD | 0x08); 
	_delay_ms(1);	                      //延遲時間
	PORTB = (PORTB & 0xC0) | (b & 0x3F); 
	PORTD = (PORTD & 0x3F) | (b & 0xC0); 
	_delay_ms(1);	
	PORTD = (PORTD & 0xF7); 
	_delay_ms(1);
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

void LCD_Lead()           //LCD 前置作業 ps:一定要打
{
	_delay_ms(15);       //延遲時間
	LCD_instruction(0,0,0x30); //功能設定
	_delay_ms(5);        //延遲時間
	LCD_instruction(0,0,0x30); //功能設定
	_delay_ms(1);        //延遲時間
	LCD_instruction(0,0,0x30); //功能設定
	_delay_ms(1);        //延遲時間
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


void LCD_own()
{
	LCD_airplane[0]=0x00;
	LCD_airplane[1]=0x00;
	LCD_airplane[2]=0x00;
	LCD_airplane[3]=0x1C;
	LCD_airplane[4]=0x1F;
	LCD_airplane[5]=0x00;
	LCD_airplane[6]=0x00;
	LCD_airplane[7]=0x00;
	LCD_airplane[8]=0x10;
	LCD_airplane[9]=0x0C;
	LCD_airplane[10]=0x06;
	LCD_airplane[11]=0x1F;
	LCD_airplane[12]=0x1F;
	LCD_airplane[13]=0x06;
	LCD_airplane[14]=0x0C;
	LCD_airplane[15]=0x10;
	LCD_airplane[16]=0x18;
	LCD_airplane[17]=0x18;
	LCD_airplane[18]=0x18;
	LCD_airplane[19]=0x1F;
	LCD_airplane[20]=0x1F;
	LCD_airplane[21]=0x00;
	LCD_airplane[22]=0x00;
	LCD_airplane[23]=0x00;
}


//==========================設定DDRAM的位址 功能:設定下一個要存入的資料DDRAM之位址 ==============================================================
void setup_DDRAM(uchar rD)  
{
	LCD_instruction(0,0,0x80 | rD); 
}



//==========================設定CDRAM的位址  功能:設定下一個要存入的資料CGRAM之位址 =============================================================
void LCD_DDRAM(uchar sC) 
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

LCD_own() ;
uchar a;
	for (a=0 ; a<23 ; a++)
	{
		LCD_DDRAM(0x00 + a) ;
		write_data(LCD_airplane[a]) ;
		_delay_ms(1);
		
	}
}








// Clock
void outputControl()
{
	PORTD = PORTD | 0x20; //0x20 00100000
	PORTD = PORTD & 0xDF; //0xDF 11011111
}
uchar inputKeys(void)
{
	uchar b;
	b = PINC & 0x70;
	return b;
}

//=================== 按鈕掃描 =========================================
void Button_scan(void)
{
	uchar col,btndata;
	uchar rowkey;
	for(col=0;col<4;col++)
	{
		PORTC=scan[col];
		rowkey=inputKeys();
		if(rowkey !=0x70)	
		{	
			_delay_ms(20);
			if(rowkey==0x60)  btn = btn | (0x01<<btndata);
			else if(rowkey==0x50)  btn = btn | (0x02<<btndata);
			else if(rowkey==0x30)  btn = btn | (0x04<<btndata);
			else if(rowkey==0x40)  btn = btn | (0x03<<btndata);
			else if(rowkey==0x20)  btn = btn | (0x05<<btndata);
			else if(rowkey==0x10)  btn = btn | (0x06<<btndata);
			else if(rowkey==0x00)  btn = btn | (0x07<<btndata);
			_delay_ms(20);
			btndata = btndata+3;
			
			
			
			
		}
		
		
	}
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
		*(data) = inputKeys();	// Key value
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
	uchar i;
	uchar j;

	switch(*(data))
	{
		case 0x20:
		
		case 0x21:
		
			
		case 0x22:
		
		case 0x23:			
		
			break;	
			
		case 0x24:
		
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
	DDRC = 0x0f;    //    0       0      0      0      1      1      1      1
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
	uchar LCD_count;
	Module() ;
	LCD_Lead();
	
	/* main event loop */
    while(1)
	{
		btn = 0;
		wdt_reset();
		usbPoll();
		Button_scan();	
		switch(btn)
		{
			case 0x01 :
			LCD_instruction(1,0,'1');
			_delay_ms(100);
			break;
			case 0x02 :
			LCD_instruction(1,0,'2');
			_delay_ms(100);
			break;
			case 0x04 :
			LCD_instruction(1,0,'3');
			_delay_ms(100);
			break;
			case 0x03 :
			LCD_instruction(1,0,'A');
			_delay_ms(100);
			break;
			case 0x05 :
			LCD_instruction(1,0,'B');
			_delay_ms(100);
			break;
			case 0x06 :
			LCD_instruction(1,0,'C');
			_delay_ms(100);
			break;
			case 0x07 :
			LCD_remove();
			_delay_ms(100);
			break;
			case 0x08 :
			LCD_instruction(1,0,'4');
			_delay_ms(100);
			break;
			case 0x10 :
			LCD_instruction(1,0,'5');
			_delay_ms(100);
			break;
			case 0x20 :
			LCD_instruction(1,0,'6');
			_delay_ms(100);
			break;
			case 0x18 :
			LCD_instruction(1,0,'D');
			_delay_ms(100);
			break;
			case 0x28 :
			LCD_instruction(1,0,'E');
			_delay_ms(100);
			break;
			case 0x30 :
			LCD_instruction(1,0,'F');
			_delay_ms(100);
			break;
			case 0x38 :
			LCD_remove();
			_delay_ms(100);
			break;
			case 0x40 :
			LCD_instruction(1,0,'7');
			_delay_ms(100);
			break;
			case 0x80 :
			LCD_instruction(1,0,'8');
			_delay_ms(100);
			break;
			case 0x100 :
			LCD_instruction(1,0,'9');
			_delay_ms(100);
			break;
			case 0xc0 :
			LCD_instruction(1,0,'G');
			_delay_ms(100);
			break;
			case 0x140 :
			LCD_instruction(1,0,'H');
			_delay_ms(100);
			break;
			case 0x180 :
			LCD_instruction(1,0,'I');
			_delay_ms(100);
			break;
			case 0x1c0 :
			LCD_remove();
			_delay_ms(100);
			break;
			case 0x200 :
			LCD_instruction(1,0,'*');
			_delay_ms(100);
			break;
			case 0x400 :
			LCD_instruction(1,0,'0');
			_delay_ms(100);
			break;
			case 0x800 :
			LCD_instruction(1,0,'#');
			_delay_ms(100);
			break;
			case 0x600 :
			LCD_instruction(1,0,'J');
			_delay_ms(100);
			break;
			case 0xa00 :
			LCD_instruction(1,0,'K');
			_delay_ms(100);
			break;
			case 0xc00 :
			LCD_instruction(1,0,'M');
			_delay_ms(100);
			break;
			case 0xe00 :
			LCD_remove();
			_delay_ms(100);
			break;
			default:
			break;	
		}
	}
}

/* ------------------------------------------------------------------------- */
