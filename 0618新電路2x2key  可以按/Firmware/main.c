#include <avr/io.h>
#include <avr/wdt.h>
#include <avr/interrupt.h>  /* sei():�P�त�_ */
#include <util/delay.h>     /*  _delay_ms() */
#include <avr/eeprom.h>
#include <avr/pgmspace.h>   /* usbdrv.h�һݪ��ѼƩw�q */
#include "usbdrv.h"
#include "oddebug.h"        /* �ϥΰ����������d�� */

static unsigned long timer[12];
static uchar scan[2]={0x0D,0x0E};
static uchar led_index,page;
static uchar VB_ON_OFF; //0 : off  1 : on   2 : end
static unsigned short scan_compared,scan_value,btn,keyin;
void process_keyin(unsigned short value);

// " | "�줸�B��l�A������ or �B��
// PB�MPD��X����LED   
static void outputByte(uchar b)
{
   PORTB = (PORTB & ~0x3f) | (b & 0x3f);
   PORTD = (PORTD & ~0xc0) | (b & 0xc0);

}
 
// 74LS273 �ɯ�  
static void LS273_ENABLE(uchar b)
{
	if(b==0x01)
	{
		PORTC =PORTC | 0x10;
	}
	else
	{
		PORTC =PORTC & ~0x10;//   "~" �ϦV
	}
}

//74LS244 �ɯ�
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

//�P��
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

//=================== ���s���y =========================================
unsigned short Button_scan(void)
{
	uchar col = 0,rowkey= 0,btndata=0;
	unsigned short key=0;

	for(col=0;col<2;col++)   //���y�C ��
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
/* ------------------------------ USB ���� --------------------------------- */
/* ------------------------------------------------------------------------- */

PROGMEM const char usbHidReportDescriptor[22] = {   /* USB ���i�y�z�� */
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

/* usbFunctionRead():��D���n�qUSB�˸m�n�D�j���ƮɴN�|�I�susbFunctionRead() */ 
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h��� */
uchar usbFunctionRead(uchar *data, uchar len)  //�����d �� VB 
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

/* usbFunctionWrite():��D���n�e�X�j���Ƶ�USB�˸m�ɴN�|�I�susbFunctionWrite()*/
/* �����T���i�H�i�H�Ѧ�usbdrv/usbdrv.h���*/
uchar usbFunctionWrite(uchar *data, uchar len) // VB �� �����d
{


    if(bytesRemaining == 0)
        return 1;               		/* USB�ǿ鵲�� */
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
    return bytesRemaining == 0; 		/* �p�G�o�O�̫�@���ơA�N�^��1 */
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
            return USB_NO_MSG;  /* �ϥ�usbFunctionRead()�Ӷǰe��ƨ�PC�D��*/
        }else if(rq->bRequest == USBRQ_HID_SET_REPORT){
            /* �ѩ�Ȧ��@�ӳ��i�����A�]���������iID */
            bytesRemaining = 8;
            currentAddress = 0;
            return USB_NO_MSG;  /* �ϥ� usbFunctionWrite()�qPC�D���������*/
        }else{
			/* �b���S���ϥγc��ӭn�D�A�]������ */
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
    usbDeviceDisconnect();  /* �j��s�˸m�C�|�A����榹�Ƶ{���ɡA���_�O�Q���઺ */
	
    while(--i){             /* ����USB���Y�ޥXPC�D���A���������ʧ@�A �ɶ� > 250 ms */
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        _delay_ms(1);
    }
    usbDeviceConnect();
    sei();					/* �P����餤�_ */
	_delay_ms(250);
	
	/******************* Timer0�]�w *********************/
	TIMSK |= (1 << TOIE0);
    TCCR0 |= (1 << CS02);             // set prescaler to 256 and start the timer 
    TCNT0 = 128; 
	

	
	outputByte(0x00);
	
	outputControl(0x00);
	
	
	
    while(1){               /* �D�n���ƥ�j�� */
        wdt_reset();		/* ���m�ݪ����p�ɾ� */
        usbPoll();			/* ����USB����ǿ� */
		
		keyin = Button_scan(); /* Ū��Button��T*/
		
		 
				switch(keyin)
			   {
			   // case 0x00:
			        // outputByte(0x0f);     // LED �G����}
					// outputControl(0x10);   
					// outputControl(0x00);
						
					// outputByte(0x00);	
					// outputControl(0x10);   
					// outputControl(0x00);
			   // break;
			   
			   case 0x01:
			        outputByte(0x01);     // LED �G����}
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
					
			   break;
			   
			   case 0x02:
			        outputByte(0x02);     // LED �G����}
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
			   break;
			   case 0x04:
					outputByte(0x04);     // LED �G����}
					outputControl(0x10);   
					outputControl(0x00);
					
					outputByte(0x00);	
					outputControl(0x10);   
					outputControl(0x00);
				break;
				case 0x8:
					outputByte(0x08);     // LED �G����}
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


