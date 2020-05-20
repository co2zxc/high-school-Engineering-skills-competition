#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():致能中斷 */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>
#include <avr/pgmspace.h>   /* usbdrv.h所需的參數定義 */
#include "usbdrv.h"
#include "oddebug.h"        /* 使用除錯巨集的範例 */

static unsigned long timer[12];
static uchar scan[2]={0x0D,0x0E};
static uchar led_index,page;
static uchar VB_ON_OFF; //0 : off  1 : on   2 : end
static unsigned short scan_compared,scan_value,btn,keyin;
void process_keyin(unsigned short value);

// " | "位元運算子，類似於 or 運算
// PB和PD輸出控制LED   
static void outputByte(uchar b)
{
   PORTB = (PORTB & ~0x3f) | (b & 0x3f);
   PORTD = (PORTD & ~0xc0) | (b & 0xc0);

}
 
// 74LS273 時脈  
static void LS273_ENABLE(uchar b)
{
	if(b==0x01)
	{
		PORTC =PORTC | 0x10;
	}
	else
	{
		PORTC =PORTC & ~0x10;//   "~" 反向
	}
}

//74LS244 時脈
static void LS244_ENABLE(uchar b)
{
	if(b==0x01)
	{
		PORTC =PORTC & ~0x20;
	}
	else
	{
		PORTC =PORTC | 0x20;
	}
}

static void outputControl(uchar b)
{
		PORTC = (PORTC & ~0x10) | (b & 0x10);//16
}

//致能
static void outputGreen(uchar b)
{
		PORTC = (PORTC & ~0x20) | (b & 0x20);//32
}

// static uchar inputkeys(void)
// {
	// uchar	b;
	// b = PINC & 0x0F;
   // return b;
// }

//=================== 按鈕掃描 =========================================
unsigned short Button_scan(void)
{
	uchar col = 0,rowkey= 0,btndata=0;
	unsigned short key=0;

	for(col=0;col<2;col++)   //掃描列 行
	{	
		PORTC=scan[col];
		rowkey=PINC & 0x0C; //11 00
		if(rowkey !=0x0C)
		{
			switch(rowkey)
			{
			case 0x08:     // 10 00
			    key = key | (0x01<<btndata);
			break;
			case 0x04:     // 01 00
				key = key | (0x02<<btndata);
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
	
	}
	return vb;
}
/* ------------------------------------------------------------------------- */
/* ------------------------------ USB 介面 --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {   /* USB 報告描述元 */
    0x06, 0x00, 0xff,              // USAGE_PAGE (販售商自定)
    0x09, 0x01,                    // USAGE (Vendor Usage 1)
    0xa1, 0x01,                    // COLLECTION (Application)
    0x15, 0x00,                    // LOGICAL_MINIMUM (0)
    0x26, 0xff, 0x00,              // LOGICAL_MAXIMUM (255)
    0x75, 0x08,                    // REPORT_SIZE (8)
    0x95, 0x08,                    // REPORT_COUNT (8)
    0x09, 0x00,                    // USAGE (Undefined)
    0xb2, 0x02, 0x01,              // FEATURE (Data,Var,Abs,Buf)
    0xc0                           // END_COLLECTION
};

/* 以下的變數是用來暫存目前資料傳輸的狀態 */
static uchar    currentAddress;
static uchar    bytesRemaining;

/* usbFunctionRead():當主機要從USB裝置要求大批資料時就會呼叫usbFunctionRead() */ 
/* 相關訊息可以可以參考usbdrv/usbdrv.h文件 */
uchar usbFunctionRead(uchar *data, uchar len)  //介面卡 到 VB 
{
    if(len > bytesRemaining)
		len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = scan_value;	// Key value
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

/* usbFunctionWrite():當主機要送出大批資料給USB裝置時就會呼叫usbFunctionWrite()*/
/* 相關訊息可以可以參考usbdrv/usbdrv.h文件*/
uchar usbFunctionWrite(uchar *data, uchar len) // VB 到 介面卡
{


    if(bytesRemaining == 0)
        return 1;               		/* USB傳輸結束 */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		
		switch(*(data)){
		case 0x05:
		
		
		
		break;

		
		}
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; 		/* 如果這是最後一批資料，就回傳1 */
}

/* USB HID群組要求: GET_REPORT與SET_REPORT */
usbMsgLen_t usbFunctionSetup(uchar data[8])
{
	usbRequest_t    *rq = (void *)data;
    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* 由於僅有一個報告類型，因此忽略報告ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* 使用usbFunctionRead()來傳送資料到PC主機*/
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* 由於僅有一個報告類型，因此忽略報告ID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* 使用 usbFunctionWrite()從PC主機接收資料*/
        }else{
			/* 在此沒有使用販售商要求，因此忽略 */
		}
    }
    return 0;
}

/* ------------------------------------------------------------------------- */
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
	uchar i = 0;
    wdt_enable(WDTO_1S);
					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3f;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
	PORTB = 0x0;    //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------30 7f
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x33;    //    0      0      1      1      0      0     1     1  
					//  (Read)  reset    out    out    out    out    out    out 						 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
    PORTC = 0x7f;    //    1      1      1     1      1      1      1      1
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xeB;    //    1      1      1      0      1      0      1      1    
					//   out    out    out    D-     out    D+     out    out
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
	
    usbInit();
    usbDeviceDisconnect();  /* 強制重新裝置列舉，當執行此副程式時，中斷是被除能的 */
	
    while(--i){             /* 模擬USB插頭拔出PC主機，假的脫離動作， 時間 > 250 ms */
        wdt_reset();		/* 重置看門狗計時器 */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();					/* 致能整體中斷 */
	_delay_ms(250);
	
	/******************* Timer0設定 *********************/
	TIMSK |= (1 << TOIE0);
    TCCR0 |= (1 << CS02);             // set prescaler to 256 and start the timer 
    TCNT0 = 128; 
	

	
	outputByte(0x00);
	
	outputControl(0x00);
	
	
	
    while(1){               /* 主要的事件迴圈 */
        wdt_reset();		/* 重置看門狗計時器 */
        usbPoll();			/* 輪詢USB控制傳輸 */
		
		keyin = Button_scan(); /* 讀取Button資訊*/
		
		 
				switch(keyin)
			   {
			   // case 0x00:
			        // outputByte(0x0f);     // LED 亮的位址
					// outputControl(0x10);   
					// outputControl(0x00);
						
					// outputByte(0x00);	
					// outputControl(0x10);   
					// outputControl(0x00);
			   // break;
			   
			   case 0x01:
			        outputByte(0x01);     // LED 亮的位址
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
					
			   break;
			   
			   case 0x02:
			        outputByte(0x02);     // LED 亮的位址
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
			   break;
			   case 0x04:
					outputByte(0x04);     // LED 亮的位址
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
				break;
				case 0x8:
					outputByte(0x08);     // LED 亮的位址
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
				break;
			   
			   // default:
               // break;
			   
			   }

			}

		
        return 0;
}


