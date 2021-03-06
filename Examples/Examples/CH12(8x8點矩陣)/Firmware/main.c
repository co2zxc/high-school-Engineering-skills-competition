/*   全華圖書: USB介面設計與應用設計入門韌體範例程式-CH12(許永和編著)                                */
/*   程式功能: 利用PB或PD驅動8顆LED，以及透過PC0~PC5、PD1與PD3驅動4顆8x8點矩陣，和使用PD0讀取按鍵值  */   

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():致能中斷 */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>

#include <avr/pgmspace.h>   /* usbdrv.h所需的參數定義 */
#include "usbdrv.h"
#include "oddebug.h"        /* 使用除錯巨集的範例 */

static uchar keyin;
static uchar LED[4][8];
/* 周邊埠PC、PD1與PD3輸出控制8*8點矩陣 */
static void outputByte(uchar b)	
{
    PORTC = (PORTC & ~0x3f) | (b      & 0x3f);
    PORTD = (PORTD & ~0x02) | (b >> 5 & 0x02);
	PORTD = (PORTD & ~0x08) | (b >> 4 & 0x08);
}
/* 周邊埠PB與PD輸出控制LED */
static void outputLED(uchar b)
{
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
}

// 74LS273 時脈
static void outputControl(uchar b)
{
	PORTD = (PORTD & ~0x20) | (b & 0x20);
}
// 透過PD0讀取按鍵值
static uchar inputKeys()
{
	uchar b;
	b = PIND & 0x01;
	return b;
}
/* 8x8點矩陣控制函式 */
static void Matrix(uchar Row, uchar Data)
{
	uchar i;
	for(i = 0 ; i < 8 ; i++)
	{
		wdt_reset();
		usbPoll();
		keyin = inputKeys();
		outputByte(Row + i);
		PORTC = (PORTC & ~0x40) | (0x40);
		//_delay_ms(1);
		PORTC = (PORTC & ~0x40) | (0x00);		
		outputByte(LED[Data][i]);
		_delay_ms(1);
	}
}

/* ------------------------------------------------------------------------- */
/* ------------------------------ USB 介面 --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {     /* USB 報告描述元 */
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
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data) = keyin;	// 按鍵值
		*(data+1) = 0;		// 保留 (0)
		*(data+2) = 0;		// 保留 (0)
		*(data+3) = 0;		// 保留 (0)
		*(data+4) = 0;		// 保留 (0)
		*(data+5) = 0;		// 保留 (0)
		*(data+6) = 0;		// 保留 (0)
		*(data+7) = 0;		// 保留 (0)
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
        return 1;               	/* USB傳輸結束 */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		
		if(*(data) == 0x20) {
			LED[data[1]][data[2]] = data[3];
		} 
		else if(*(data) == 0x21) 
		{
			outputLED(data[1]);
			outputControl(0x20);	// ---- 產生一個時脈
			outputControl(0x0);		
		}
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0;		/* 如果這是最後一批資料，就回傳1 */
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
	DDRC = 0x7F;    //    0      1      1      1      1      1      1      1  
					//  (Read)  out    out    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
    PORTC = 0x40;   //    0      1      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xEA;    //    1      1      1      0      1      0      1      0    
					//   out    out    out    D-     out    D+     out    in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
	
    usbInit();
    usbDeviceDisconnect();  /* 強制重新裝置列舉，當執行此副程式時，中斷是被除能的 */
	i = 0;
    while(--i){             /* 模擬USB插頭拔出PC主機，假的脫離動作， 時間 > 250 ms */
        wdt_reset();		/* 重置看門狗計時器 */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();					/* 致能整體中斷 */
    while(1){               /* 主要的事件迴圈 */ 
        wdt_reset();		/* 重置看門狗計時器 */
        usbPoll();			/* 輪詢USB控制傳輸 */
		Matrix(0x08,0);		/* 改變8x8點矩陣U3-1資訊 */
		Matrix(0x10,1);		/* 改變8x8點矩陣U3-2資訊 */
		Matrix(0x20,2);		/* 改變8x8點矩陣U3-3資訊 */
		Matrix(0x40,3);		/* 改變8x8點矩陣U3-4資訊 */
    }
    return 0;
}