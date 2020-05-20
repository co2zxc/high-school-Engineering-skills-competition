#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* for sei() */
#include <util/delay.h>     /* for _delay_ms() */
#include <avr/eeprom.h>
#include <avr/pgmspace.h>   /* required by usbdrv.h */
#include "usbdrv.h"
#include "oddebug.h"        /* This is also an example for using debug macros */

#define KeyCol1	!(PIND&0x01)	/* PD0 */
#define KeyCol2	!(PIND&0x02)	/* PD1 */
#define KeyCol3	!(PIND&0x08)	/* PD3 */
//#define FO(b) PORTD = (PORTD & ~0x20) | ((b<<5) & 0x20);	// PD5
#define RS(b) PORTC = (PORTC & ~0x10) | ((b<<4) & 0x10);	// PC4
#define  E(b) PORTC = (PORTC & ~0x20) | ((b<<5) & 0x20);	// PC5
#define read_eeprom_byte(address) eeprom_read_byte ((const uint8_t*)address)
#define write_eeprom_byte(address,value) eeprom_write_byte ((uint8_t*)address,(uint8_t)value)

uchar keyin;
uchar ClockTimeData[6] = {0x00,0x00,0x00,0x00,0x00,0x00};
uchar AlarmTimeData[6] = {0x30,0x30,0x30,0x30,0x30,0x30};
uchar mode = 0;
/* 周邊埠PB與PD輸出控制LCD所顯示的資料 */
static void LCD_Control(uchar b)
{
    PORTB = (PORTB & ~0x3f) | (b & 0x3f);
    PORTD = (PORTD & ~0xc0) | (b & 0xc0);
}
/* LCD 控制 */
void LCDSet(uchar LCDdata , uchar data)
{ 
	LCD_Control(LCDdata);
	RS(data);
	E(1);
	_delay_ms(1);
	E(0);
	_delay_ms(1);
}  
/* LCD 初始化 */
void LCDInit(void)
{
	LCDSet(0x38,0); //選擇功能
	LCDSet(0x1f,0); //設定游標移動方式
	LCDSet(0x01,0); //清除顯示器
//	LCDSet(0x0f,0); //開啟顯示器
	LCDSet(0x0c,0); //開啟顯示器(不顯示游標) 
} 
/* 掃描4*3鍵盤 */
uchar scankey(void)
{
	uchar i = 0,key = 0,scan = 0x01;
	PORTC=~scan;
	for(i=0;i<4;i++)
	{
		if(KeyCol1) break;
		key++;
		if(KeyCol2) break;
		key++;
		if(KeyCol3) break;
		key++;
		PORTC=~(scan<<=1);
	}
	return key;
}
/* 讀取4*3鍵盤 */
uchar getkey(void)
{
	uchar key;
	key=scankey(); /* 掃瞄4*3鍵盤 */
	_delay_ms(10);
	return key;
}
/* LCD時間顯示 */
void LCDTimePrint(void)
{
	LCDSet(0x84,0);
	LCDSet(ClockTimeData[0],1);
	LCDSet(ClockTimeData[1],1);
	LCDSet(0x3A,1);
	LCDSet(ClockTimeData[2],1);
	LCDSet(ClockTimeData[3],1);
	LCDSet(0x3A,1);
	LCDSet(ClockTimeData[4],1);
	LCDSet(ClockTimeData[5],1);
}
/* LCD鬧鐘顯示 */
void LCDAlarmTimePrint(void)
{
	LCDSet(0x84,0);
	LCDSet(AlarmTimeData[0],1);
	LCDSet(AlarmTimeData[1],1);
	LCDSet(0x3A,1);
	LCDSet(AlarmTimeData[2],1);
	LCDSet(AlarmTimeData[3],1);
	LCDSet(0x3A,1);
	LCDSet(AlarmTimeData[4],1);
	LCDSet(AlarmTimeData[5],1);
}
/* LCD 顯示 */
void LCDPrint(uchar *LCDdata)
{
	unsigned int i;
	for(i=0;LCDdata[i]!='\0';i++)
	{
		if(i==16) LCDSet(0xc0,0) ;  //換行
		LCDSet(LCDdata[i],1) ;
	}
}
/*
void Fout(unsigned int f) // 頻率輸出 f的單位:Hz
{
	FO(1);
	_delay_us(500000/f);
	FO(0);
	_delay_us(500000/f);	
}
*/

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
    0x95, 0x08,                    // REPORT_COUNT (128)
    0x09, 0x00,                    // USAGE (Undefined)
    0xb2, 0x02, 0x01,              // FEATURE (Data,Var,Abs,Buf)
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
	if(currentAddress == 0){
		*(data) = read_eeprom_byte(0);			// Reserve (zero)
		*(data+1) = read_eeprom_byte(1);		// Reserve (zero)
		*(data+2) = read_eeprom_byte(2);		// Reserve (zero)
		*(data+3) = read_eeprom_byte(3);		// Reserve (zero)
		*(data+4) = read_eeprom_byte(4);		// Reserve (zero)
		*(data+5) = read_eeprom_byte(5);		// Reserve (zero)
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
		if(*(data) == 0x20) {
			ClockTimeData[0] = data[1] + 0x30;
			ClockTimeData[1] = data[2] + 0x30;
			ClockTimeData[2] = data[3] + 0x30;
			ClockTimeData[3] = data[4] + 0x30;
			ClockTimeData[4] = data[5] + 0x30;
			ClockTimeData[5] = data[6] + 0x30;
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

int main(void)
{	
	uchar i = 0, b = 0;
	uchar key = 0, Location = 0, bingo = 0;
    wdt_enable(WDTO_1S);
					// ---------------------------------------------------------
	                //   DDB7   DDB6   DDB5   DDB4   DDB3   DDB2   DDB1   DDB0
	DDRB = 0x3f;    //    0      0      1      1      1      1      1      1
	                //    x      x     out    out    out    out    out    out
					// ---------------------------------------------------------
                    //  PORTB7 PORTB6 PORTB5 PORTB4 PORTB3 PORTB2 PORTB1 PORTB0
//	PORTB = 0x00;   //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

					// ---------------------------------------------------------
				    //    -     DDC6   DDC5   DDC4   DDC3   DDC2   DDC1   DDC0
	DDRC = 0x3f;    //    0      0      1      1      1      1      1      1  
					//  (Read)  in     out    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
//  PORTC = 0x40;   //    0      1      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xc0;    //    1      1      0      0      0      0      0      0    
					//   out    out     in     D-     in    D+      in     in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
//	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
	
    usbInit();
    usbDeviceDisconnect();  /* enforce re-enumeration, do this while interrupts are disabled! */
	i = 0;
    while(--i){             /* fake USB disconnect for > 250 ms */
		wdt_reset();
		_delay_ms(1);
    }
    usbDeviceConnect();
    sei();
	LCDInit(); 					/* LCD 初始化     */
	for(i=0;i<6;i++) AlarmTimeData[i]=read_eeprom_byte(i);
    while(1)
	{                /* main event loop */
	    wdt_reset();
        usbPoll();				/* 輪詢USB控制傳輸 */
		key = getkey() + 0x31;
		if(key == 0x3B) key = 0x30;
		else if(key == 0x3A) key = 0x2A;
		else if(key == 0x3C) key = 0x23;

		if(key == '*') mode = ~mode;
		if(keyin != key)
		{
			if(mode) 
			{
				b = 0;
				bingo = 0;
				if(key <= '9' && key >= '0'){
					AlarmTimeData[Location] = key;
					write_eeprom_byte(Location,key);
					if(Location<=5) Location++;  else Location=6;
				}
				if(key == '#') mode = 0;
			}	
			keyin = key;
		}
		if(!mode) 
		{
			bingo = 0;
			Location = 0;
			for(i = 0;i < 6;i++) 
				if(ClockTimeData[i] == AlarmTimeData[i]) bingo++;
			LCDTimePrint();
			LCDSet(0xC3,0);
			if(bingo == 6) b = 1;  
			if(b) LCDPrint("  Time Up ");
			else  LCDPrint("Clock Mode");
		}
		else 
		{
			LCDAlarmTimePrint();
			LCDSet(0xC3,0);
			LCDPrint("Alarm Mode");
		}
    }
    return 0;
}
