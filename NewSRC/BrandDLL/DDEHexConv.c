#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <windows.h>

//FILE *fpd; 
int _stdcall ShttoHex( short ,char * );
int _stdcall InttoHex( int, char * );
int _stdcall FlttoHex( float, char * );
int _stdcall DbltoHex( double, char * );
int _stdcall HextoSht( char *, short * );
int _stdcall HextoInt( char *, int * );
int _stdcall HextoFlt( char *, float * );
int _stdcall HextoDbl( char *, double * );
unsigned int ChrtoHex (  char );
char HextoChr ( int );

/*							
int _stdcall ShttoHex( short ,BSTR );
int _stdcall InttoHex( int, BSTR );
int _stdcall FlttoHex( float, BSTR );
int _stdcall DbltoHex( double, BSTR );
int _stdcall HextoSht( BSTR, short * );
int _stdcall HextoInt( BSTR, int * );
int _stdcall HextoFlt( BSTR, float * );
int _stdcall HextoDbl( BSTR, double * );
unsigned int ChrtoHex (  char );
char HextoChr ( int );
*/

/* < ShttoHex > *********************************************
	short変数(2byte) -> ヘキサ文字(4byte)
**************************************************************/
//int _stdcall ShttoHex( short insht ,char hchr[] )
int _stdcall ShttoHex( short insht ,char *hchr )
{

	int i,wc,j=0 ,kos,kos2;

	//if((fpd = fopen("debug.dat","w")) == NULL) return 1;

	kos  = sizeof(short)*2;

	/* 4bit shift */
	
	kos2 = kos*4 - 4 ;
	for ( i=kos2;i>=0;i-=4) {
		wc = (insht>>i)&0x0f ;
		if ( wc<0 || wc>15 ) return(1);
		hchr[j] = HextoChr( wc );
		j++;
	}
	if ( j!=kos ) return(1);
	hchr[4] = '\0';

	//fprintf(fpd,"hchr=%s\n",hchr);

	return(0);
}
/* < InttoHex > *********************************************
	int変数(4byte) -> ヘキサ文字(8byte)
**************************************************************/
int _stdcall InttoHex( int inint , char *hchr)
{

	int i ,wc,j=0,kos,kos2;

	kos  = sizeof(int)*2;
	
	/* 4bit shift */
	kos2 = kos*4 - 4 ;
	for ( i=kos2;i>=0;i-=4) {
		wc = (inint>>i)&0x0f ;
		if ( wc<0 || wc>15 ) return(1);
		hchr[j] = HextoChr( wc );
		j++;
	}

	if ( j!=kos ) return(1);
	hchr[j] = '\0';

	return(0);
}	
/* < FlttoHex > *********************************************
	float変数(4byte) -> ヘキサ文字(8byte)
**************************************************************/
int _stdcall FlttoHex( float inflt, char *hchr )
{

	int i ,wc,j=0,kos,kos2;
	unsigned int uibuf;

	kos  = sizeof(float)*2;
	
	/* Int Area Copy */
	memcpy(&uibuf,&inflt,sizeof(float) );

	/* 4bit shift */
	kos2 = kos*4 - 4 ;
	for ( i=kos2;i>=0;i-=4) {
		wc = (uibuf>>i)&0x0f ;
		if ( wc<0 || wc>15 ) return(1);
		hchr[j] = HextoChr( wc );
		j++;
	}

	if ( j!=kos ) return(1);
	
	hchr[j] = '\0';

	return(0);
}
/* < DbltoHex > *********************************************
	double変数(8byte) -> ヘキサ文字(16byte)
**************************************************************/
int _stdcall DbltoHex( double indbl, char *hchr )
{

	int i ,wc,j=0,kos,kos2;
	unsigned int uibuf[2];

	kos=sizeof(double)*2 ;

	/* Shift buffer Copy Int Area */
	memcpy(&uibuf,&indbl,sizeof(double) );

	/* 4bit shift UP 4byte */
	kos2 = kos*2 - 4 ;
	for ( i=kos2;i>=0;i-=4) {
		wc = (uibuf[0]>>i)&0x0f ;
		if ( wc<0 || wc>15 ) return(1);
		hchr[j] = HextoChr( wc );
		j++;
	}
	/* 4bit shift Down 4byte */
	for ( i=kos2;i>=0;i-=4) {
		wc = (uibuf[1]>>i)&0x0f ;
		if ( wc<0 || wc>15 ) return(1);
		hchr[j] = HextoChr( wc );
		j++;
	}

	if ( j!=kos ) return(1);

	hchr[j] = '\0';

	return(0);
}	

/* < HextoSht > *********************************************
	ヘキサ文字(4byte) -> short変数(2byte)  
**************************************************************/
int _stdcall HextoSht( char *hchr, short *outsht )
{

	int i ,kos;
	unsigned int uibuf=0 ;

	kos=sizeof(short)*2 ;

	uibuf =  ChrtoHex( hchr[0] ) ;
	for (i=1;i<kos;i++) { 
		uibuf = (uibuf<<4) + ChrtoHex( hchr[i] ) ;
	}
	*outsht = (short)uibuf ;
	return(0);
}
/* < HextoInt > *********************************************
	ヘキサ文字(8byte) ->  int変数(4byte)
**************************************************************/
int _stdcall HextoInt( char *hchr, int *outint )
{

	int i ,kos;
	unsigned int uibuf=0 ;

	kos=sizeof(int)*2 ;

	uibuf =  ChrtoHex( hchr[0] ) ;
	for (i=1;i<kos;i++) { 
		uibuf = (uibuf<<4) + ChrtoHex( hchr[i] ) ;
	}

	*outint = uibuf ;
	return(0);
}	
/* < HextoFlt > *********************************************
	ヘキサ文字(8byte) -> float変数(4byte) 
**************************************************************/
int _stdcall HextoFlt( char *hchr, float *outflt )
{

	int i ,kos;
	unsigned int uibuf=0 ; 

	kos=sizeof(float)*2 ;

	uibuf =  ChrtoHex( hchr[0] ) ;
	for (i=1;i<kos;i++) { 
		uibuf = (uibuf<<4) + ChrtoHex( hchr[i] ) ;
	}

	memcpy(outflt,&uibuf,sizeof(float) );

	return(0);
}
/* < HextoDbl > *********************************************
	ヘキサ文字(16byte)  -> double変数(8byte) 
**************************************************************/
int _stdcall HextoDbl( char *hchr, double *outdbl )
{

	int i;
	unsigned int uibuf[2] ;
	
	uibuf[0] = 0 ; 
	uibuf[1] = 0 ; 

	uibuf[0] =  ChrtoHex( hchr[0] ) ;
	for (i=1;i<8;i++) { 
		uibuf[0] = (uibuf[0]<<4) + ChrtoHex( hchr[i] ) ;
	}

	uibuf[1] =  ChrtoHex( hchr[8] ) ;
	for (i=9;i<16;i++) { 
		uibuf[1] = (uibuf[1]<<4) + ChrtoHex( hchr[i] ) ;
	}

	memcpy(outdbl,&uibuf[0],8);

	return(0);
}	

/* < ChrtoHex > *********************************************
	ヘキサ文字(1byte)  -> unsigned int  
**************************************************************/

unsigned int ChrtoHex (  char bitdat )
{
char    CHEX[16] = { '0','1','2','3','4','5','6','7',
		  			 '8','9','A','B','C','D','E','F'  };

	int   j;

	for(j=0;j<16;j++) {
		if ( !_strnicmp( &bitdat,&CHEX[j],1 ) ) {
			 	return(j) ;
		} 
	} 	

	return(0);
}
/* < ChrtoHex > *********************************************
	int -> ヘキサ文字(1byte)  
**************************************************************/

char HextoChr (  int bitdat )
{
char    CHEX[16] = { '0','1','2','3','4','5','6','7',
		  			 '8','9','A','B','C','D','E','F'  };


 	return( CHEX[bitdat]) ;
}
