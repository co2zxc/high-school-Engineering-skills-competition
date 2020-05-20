/*   �{���\��: �N 4*3 �������ܦbLCD�W   */

#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():�P�त�_ */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>
#include <avr/pgmspace.h>   /* usbdrv.h�һݪ��ѼƩw�q */
#include "usbdrv.h"
#include "oddebug.h"        /* �ϥΰ����������d�� */

#define KeyCol1	!(PINC&0x20)	/* PC5 */
#define KeyCol2	!(PIND&0x01)	/* PD0 */
#define KeyCol3	!(PIND&0x02)	/* PD1 */
#define RS(b) PORTD = (PORTD & ~0x08) | ((b<<3) & 0x08);	// PD3
#define  E(b) PORTD = (PORTD & ~0x20) | ((b<<5) & 0x20);	// PD5
//#define RW(b)  PORTC = (PORTC & ~0x10)  |  (b<<4);			// PC4

static uchar keyin, OutData = 0x20;
unsigned char ClockTimeData[6] = {0x32,0x33,0x35,0x39,0x35,0x39};
unsigned char AlarmTimeData[6] = {0x30,0x30,0x30,0x30,0x30,0x30};
unsigned char mode = 0,  b = 0;
/* �P���PB�PPD��X����LCD����ܪ���� */
static void LCD_Control(unsigned char b)
{
    PORTB = (PORTB & ~0x3f) | (b & 0x3f);
    PORTD = (PORTD & ~0xc0) | (b & 0xc0);
}
/* LCD ���� */
void LCDSet(unsigned char LCDdata , unsigned char data)
{ 
	LCD_Control(LCDdata);
	RS(data);
	E(1);
	_delay_ms(1);
	E(0);
	_delay_ms(1);
}  
/* LCD ��l�� */
void LCDInit()
{
	LCDSet(0x38,0); //��ܥ\��
	LCDSet(0x1f,0); //�]�w��в��ʤ覡
	LCDSet(0x01,0); //�M����ܾ�
	LCDSet(0x0f,0); //�}����ܾ� 
} 
/* ���y4*3��L */
unsigned char scankey()
{
	unsigned char i = 0,key = 0,scan = 0x01;
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
/* Ū��4*3��L */
unsigned char getkey()
{
	unsigned char key;
	key=scankey(); /* ����4*3��L */
	_delay_ms(1);
	return key;
}
/* LCD�ɶ���� */
void LCDTimePrint()
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
/* LCD�x����� */
void LCDAlarmTimePrint()
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
/* LCD ��� */
void LCDPrint(unsigned char *LCDdata)
{
	unsigned int i;
	for(i=0;LCDdata[i]!='\0';i++)
	{
		if(i==16) LCDSet(0xc0,0) ;  //����
		LCDSet(LCDdata[i],1) ;
	}
}
/* ------------------------------------------------------------------------- */
/* ------------------------------ USB ���� --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {    /* USB ���i�y�z�� */
    0x06, 0x00, 0xff,              // USAGE_PAGE (�c��Ӧ۩w)
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

/* �H�U���ܼƬO�ΨӼȦs�ثe��ƶǿ骺���A */
static uchar    currentAddress;
static uchar    bytesRemaining;

/* usbFunctionRead():��D���n�qUSB�˸m�n�D�j���ƮɴN�|�I�susbFunctionRead()*/ 
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h��� */
uchar   usbFunctionRead(uchar *data, uchar len)
{
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		*(data)   = b;							// �����
		*(data+1) = AlarmTimeData[0] - 0x30;	// ��-�Q��� (0)
		*(data+2) = AlarmTimeData[1] - 0x30;	// ��-�Ӧ�� (0)
		*(data+3) = AlarmTimeData[2] - 0x30;	// ��-�Q��� (0)
		*(data+4) = AlarmTimeData[3] - 0x30;	// ��-�Ӧ�� (0)
		*(data+5) = AlarmTimeData[4] - 0x30;	// ��-�Q��� (0)
		*(data+6) = AlarmTimeData[5] - 0x30;	// ��-�Ӧ�� (0)
		*(data+7) = mode;						// Out���}��� (0)
	}
    currentAddress += len;
    bytesRemaining -= len;
    return len;
}

/* usbFunctionWrite():��D���n�e�X�j���Ƶ�USB�˸m�ɴN�|�I�susbFunctionWrite()*/
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h���*/
uchar   usbFunctionWrite(uchar *data, uchar len)
{
    if(bytesRemaining == 0)
        return 1;               /* USB�ǿ鵲�� */
    if(len > bytesRemaining)
        len = bytesRemaining;
	if(currentAddress == 0){
		ClockTimeData[0] = (data[0] / 10) + 0x30;
		ClockTimeData[1] = (data[0] % 10) + 0x30;
		ClockTimeData[2] = (data[1] / 10) + 0x30;
		ClockTimeData[3] = (data[1] % 10) + 0x30;
		ClockTimeData[4] = (data[2] / 10) + 0x30;
		ClockTimeData[5] = (data[2] % 10) + 0x30;
		OutData = data[3] + 0x30;
	}
    currentAddress += len;
    bytesRemaining -= len;
    return bytesRemaining == 0; /* �p�G�o�O�̫�@���ơA�N�^��1 */
}
/* USB HID�s�խn�D: GET_REPORT�PSET_REPORT */
usbMsgLen_t usbFunctionSetup(uchar data[8])
{
	usbRequest_t    *rq = (void *)data;
    if((rq->bmRequestType & USBRQ_TYPE_MASK) == USBRQ_TYPE_CLASS){    /* HID class request */
        if(rq->bRequest == USBRQ_HID_GET_REPORT){  /* wValue: ReportType (highbyte), ReportID (lowbyte) */
            /* �ѩ�Ȧ��@�ӳ��i�����A�]���������iID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* �ϥ�usbFunctionRead()�Ӷǰe��ƨ�PC�D�� */
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* �ѩ�Ȧ��@�ӳ��i�����A�]���������iID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* �ϥ� usbFunctionWrite()�qPC�D��������� */
        }
    }else{
        /* �b���S���ϥγc��ӭn�D�A�]������ */
    }
    return 0;
}

int main(void)
{	
	uchar i = 0;
	uchar key = 0, Location = 0, bingo = 0;
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
	DDRC = 0x5F;    //    0      1      0      1      1      1      1      1  
					//  (Read)  out     in    out    out    out    out    out							 
					// ---------------------------------------------------------
                    //    -    PORTC6 PORTC5 PORTC4 PORTC3 PORTC2 PORTC1 PORTC0							 
    PORTC = 0x0;    //    0      0      0      0      0      0      0      0
					// ---------------------------------------------------------

    				// ---------------------------------------------------------
	                //   DDD7   DDD6   DDD5   DDD4   DDD3   DDD2   DDD1   DDD0
	DDRD = 0xE8;    //    1      1      1      0      1      0      0      0    
					//   out    out    out     D-    out     D+     in     in
					// ---------------------------------------------------------
	                //  PORTD7 PORTD6 PORTD5 PORTD4 PORTD3 PORTD2 PORTD1 PORTD0
	PORTD = 0x00;   //    0      0      0      0      0      0      0      0  
					// ---------------------------------------------------------
    usbInit();
    usbDeviceDisconnect();  	/* �j��s�˸m�C�|�A����榹�Ƶ{���ɡA���_�O�Q���઺ */
	
    while(--i){            		/* ����USB���Y�ޥXPC�D���A���������ʧ@�A �ɶ� > 250 ms */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();						/* �P����餤�_ */
	LCDInit(); 					/* LCD ��l��     */
    while(1)					/* �D�n���ƥ�j�� */
	{                
        usbPoll();				/* ����USB����ǿ� */
		key = getkey() + 0x31;
		if(key == 0x3B) key = 0x30;
		else if(key == 0x3A) key = 0x2A;
		else if(key == 0x3C) key = 0x23;
		
		LCDSet(0x80,0);
		LCDSet(OutData,1);

		if(keyin != key)
		{
			if(key == '*') mode = ~mode;
			if(mode) 
			{
				b = 0;
				bingo = 0;
				if(key <= '9' && key >= '0') AlarmTimeData[Location++] = key;
				if(key == '#') mode = 0;
				keyin = key;
			}	
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
			if(b) 
			{
				LCDPrint("  Time Up ");
				//for(i = 0; i < 6; i++)  AlarmTimeData[i] = 0x30;
			}
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
