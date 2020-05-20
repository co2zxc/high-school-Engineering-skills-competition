/*   全華圖書: USB介面設計與應用設計入門韌體範例程式-CH13(許永和編著) */
/*   程式功能: 利用PB或PD驅動4顆七段顯示器                             */   

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():致能中斷 */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* usbdrv.h所需的參數定義 */
#include "usbdrv.h"
#include "oddebug.h"        /* 使用除錯巨集的範例 */
/* 七段顯示器 0~F數值的解碼數值 */ 
/* PB解碼數值 */
unsigned char Led_B[16] = {0x3F,0x06,0x1B,0x0F,0x26,0x2D,0x3D,0x27,0x3F,0x27,0x37,0x3C,0x39,0x1E,0x39,0x31};
/* PD解碼數值 */
unsigned char Led_D[16] = {0x00,0x00,0x40,0x40,0x40,0x40,0x40,0x00,0x40,0x40,0x40,0x40,0x00,0x40,0x40,0x40};
unsigned char SEG_Data[4] = {0,0,0,0};
/* 周邊埠PB與PD輸出控制七段顯示器所顯示的資料 */
static void SEG_Control(uchar b)
{
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
}
/* 七段顯示器函式，四顆七段顯示器(U4，U5，U6與U7)個別控制 */
static void SEG(uchar w, uchar x, uchar y, uchar z)
{
	PORTD = ((PORTD & ~0x2B) | 0x01);		/* U4 */
	SEG_Control(Led_B[w]|Led_D[w]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x02);		/* U5 */
	SEG_Control(Led_B[x]|Led_D[x]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x08);		/* U6 */
	SEG_Control(Led_B[y]|Led_D[y]);
	_delay_ms(1);
	PORTD = ((PORTD & ~0x2B) | 0x20);		/* U7 */
	SEG_Control(Led_B[z]|Led_D[z]);
	_delay_ms(1);
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

/* usbFunctionRead():當主機要從USB裝置要求大批資料時就會呼叫usbFunctionRead()*/ 
/* 相關訊息可以可以參考usbdrv/usbdrv.h文件*/
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
	}
    currentAddress += len;
    bytesRemaining -= len;
    return len;
}

/* usbFunctionWrite():當主機要送出大批資料給USB裝置時就會呼叫usbFunctionWrite()*/
/* 相關訊息可以可以參考usbdrv/usbdrv.h文件*/
 
uchar   usbFunctionWrite(uchar *data, uchar len)
{
    if(bytesRemaining == 0)
        return 1;               /* USB傳輸結束 */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0)
	{	
		/* 輸出與推動SEG的資料位元組 */
		SEG_Data[0] = data[0];	
		SEG_Data[1] = data[1];
		SEG_Data[2] = data[2];
		SEG_Data[3] = data[3];
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* 如果這是最後一批資料，就回傳1 */
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
        }
    }else{
        /* 在此沒有使用販售商要求，因此忽略 */
    }
    return 0;
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

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x7f;    //    0      1      1      1      1      1      1      1  
					//  (Read)  out    out    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
    PORTC = 0x0;    //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xeb;    //    1      1      1      0      1      0      1      1    
					//   out    out    out    D-     out     D+     out    out
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
    while(1)				/* 主要的事件迴圈 */
	{                
        wdt_reset();		/* 重置看門狗計時器 */
        usbPoll();			/* 輪詢USB控制傳輸 */
		SEG(SEG_Data[0],SEG_Data[1],SEG_Data[2],SEG_Data[3]);	/* 改變SEG資訊 */
    }
    return 0;
}