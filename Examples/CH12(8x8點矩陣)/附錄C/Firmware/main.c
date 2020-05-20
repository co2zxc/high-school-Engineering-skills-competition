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
uchar LOGO=0,z=0,program_run=0;
/*  周邊埠PB與PD輸出至74273(LED)與74244(點矩陣)  */
static void outputByte(uchar b)
{
	PORTB = (PORTB & ~0x3F) | (b & 0x3F);
	PORTD = (PORTD & ~0xC0) | (b & 0xC0);
}
static void outputs(uchar b)
{
	PORTC = (PORTC & ~0x7F) | (b & 0x7F);
}
// 74273 時脈
static void output273(uchar b)
{
	PORTD = (PORTD & ~0x20) | (b & 0x20);
}

// 透過PD0讀取按鍵值
static uchar inputKeys()
{
	uchar b;
	b = PIND & 0x03;
	return b;
}

// 透過PD1讀取按鍵值
/* 8x8點矩陣控制函式 */
static void Matrix(uchar Row, uchar Data)
{
	uchar i;
	for(i = 0 ; i < 8 ; i++)
	{
		wdt_reset();
		usbPoll();
		keyin = inputKeys();
		outputs(Row+i);	
		outputByte(LED[Data][i]);
		_delay_ms(1);
	}
		
}
static void Matrixbf(uchar Row, uchar Data)
{
	uchar num[20][8] = {{0x00,0xc2,0xa1,0x91,0x89,0x86,0x00,0x00},  		
					{0x41,0x81,0x89,0x95,0x63,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x4f,0x89,0x89,0x89,0x71},			
					{0x00,0x00,0x06,0x89,0x89,0x49,0x3e,0x00},			//23:59
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x38,0x24,0x22,0xff,0x20},			
					{0x00,0x00,0x01,0x01,0xf9,0x05,0x03,0x00},			//  :47
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},		
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},			
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00},//   00
					
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},			
					{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x66},			
					{0x66,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},			
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00},    // :00
					
					{0x00,0x7e,0x81,0x81,0x81,0x7e,0x00,0x00},
					{0x7e,0x81,0x81,0x81,0x7e,0x00,0x00,0x66},
					{0x66,0x00,0x00,0x7e,0x81,0x81,0x81,0x7e},
					{0x00,0x00,0x7e,0x81,0x81,0x81,0x7e,0x00}};   // 00:00
								
					
					
	uchar pic[64] = {0x3f,0xa1,0xe1,0xe1,0xe1,0xe1,0xa1,0x3f,
					0x00,0x06,0x05,0x09,0x91,0xe1,0xe1,0x91,
					0x09,0x05,0x06,0x00,0xc0,0xc0,0xf8,0x88,
					0x8f,0x81,0x8f,0x88,0xf8,0xc0,0xc0,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
					0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00};
	uchar i;
	
	if(keyin==0x00)  
	{
		for(i = 0 ; i < 8 ; i++)
		{
			wdt_reset();
			usbPoll();
			keyin = inputKeys();
			outputs(Row+i);	
			outputByte(pic[Data*8+i+LOGO]);
			_delay_ms(1);
		}		
	}
	else
	{
		for(i = 0 ; i < 8 ; i++)
		{
			wdt_reset();
			usbPoll();
			keyin = inputKeys();
			outputs(Row+i);	
			outputByte(num[Data][i]);
			_delay_ms(1);
		}
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
		*(data) = keyin;		// k0 k1
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
		
		if(*(data) == 0x22) 
		{
			program_run = 1;
		}
		else if(*(data) == 0x20) {
			LED[data[1]][data[2]] = data[3];
		} 
		else if(*(data) == 0x21) 
		{
			outputByte(data[1]);
			output273(0x0);	// ---- 產生一個時脈
			output273(0x20);		
		}
		else if(*(data) == 0x23)
		{
			program_run =2;  //結束
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
    //wdt_enable(WDTO_1S);
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
    PORTC = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xE8;    //    1      1      1      0      1      0      0      0    
					//   out    out    out    D-     out    D+     in     in
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
    sei();						/* 致能整體中斷 */
	i = 0;
	outputByte(0xff);
	output273(0x0);				// ---- 產生一個時脈
	output273(0x20);		
    while(1)					// 主要的事件迴圈 */ 
	{	
		if(program_run ==2 )  //00:00
		{
			if(keyin == 0x03){
			Matrixbf(0x40,16);		// 改變8x8點矩陣U3-1資訊 
			Matrixbf(0x20,17);		// 改變8x8點矩陣U3-2資訊 
			Matrixbf(0x10,18);		// 改變8x8點矩陣U3-3資訊 
			Matrixbf(0x08,19);		// 改變8x8點矩陣U3-4資訊
			}
			if(keyin == 0x01 )		//   00
			{
				Matrixbf(0x40,8);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,9);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,10);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,11);		// 改變8x8點矩陣U3-4資訊 
			}
			if(keyin == 0x02)		//  :00
			{
				Matrixbf(0x40,12);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,13);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,14);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,15);		// 改變8x8點矩陣U3-4資訊 	
			}
			if(keyin == 0x00)
			{
				Matrixbf(0x40,4);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,4);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,4);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,4);// 改變8x8點矩陣U3-4資訊 	
			}
		}
		if(program_run == 1)
		{
			wdt_reset();		// 重置看門狗計時器 
			usbPoll();			// 輪詢USB控制傳輸 
			Matrix(0x40,0);		// 改變8x8點矩陣U3-1資訊 
			Matrix(0x20,1);		// 改變8x8點矩陣U3-2資訊 
			Matrix(0x10,2);		// 改變8x8點矩陣U3-3資訊 
			Matrix(0x08,3);		// 改變8x8點矩陣U3-4資訊 		

		}
		if(program_run == 0)
		{
			if(keyin == 0x00)	// LOGO
			{
				Matrixbf(0x40,0);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,1);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,2);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,3);		// 改變8x8點矩陣U3-4資訊 
				if(z==0)
				{
					LOGO++;
					_delay_ms(10);
					if(LOGO>31)	z=1;
				}
				if(z==1)
				{
					LOGO--;
					_delay_ms(10);
					if(LOGO<1)	z=0;
				}
			}
			if(keyin == 0x01)		//   00
			{
				Matrixbf(0x40,8);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,9);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,10);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,11);		// 改變8x8點矩陣U3-4資訊 
			}
			if(keyin == 0x02)		//  :47
			{
				Matrixbf(0x40,4);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,5);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,6);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,7);		// 改變8x8點矩陣U3-4資訊 	
			}
			if(keyin == 0x03)		//23:59
			{
				Matrixbf(0x40,0);		// 改變8x8點矩陣U3-1資訊 
				Matrixbf(0x20,1);		// 改變8x8點矩陣U3-2資訊 
				Matrixbf(0x10,2);		// 改變8x8點矩陣U3-3資訊 
				Matrixbf(0x08,3);		// 改變8x8點矩陣U3-4資訊 		
			}

		}
		
	}
		return 0;
}