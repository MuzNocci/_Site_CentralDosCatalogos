/**
 * Generic
 */

var defaultFocusOccur;
var defaultFocusField;

function selectValue( url )
{
	var hWnd = window.open( url, "PopUp", "width=250,height=325,resizable=yes,scrollbars=yes" );
	if( ( document.window != null ) && ( !hWnd.opener ) )
		hWnd.opener = document.window;
}

function updateField( field, value )
{
	field.value = value;
	window.close();
}

function isDigit( str )
{
  for( var i = 0; i < str.length; i++ )
  {
    var charCode = str.charCodeAt( i );
    if( !(charCode >= 48 && charCode <= 57) )
      return( false );
  }
  return( true );
}

function isLetter( str )
{
  for( var i = 0; i < str.length; i++ )
  {
    var charCode = str.charCodeAt( i );
    if( !(charCode >= 65 && charCode <= 90 || charCode >= 97 && charCode <= 122) )
      return( false );
  }
  return( true );
}

function isLetterOrDigit( str )
{
  for( var i = 0; i < str.length; i++ )
  {
    var character = str.charAt( i );
    if( !isDigit( character ) && !isLetter( character ) )
      return( false );
  }
  return( true );
}

function searchKey( campo, keyEvent )
{
  var key = keyEvent.keyCode;
  document.form[campo].value = key;
}


function nextField( field, event )
{
  var size = field.maxLength;
  if (isNaN(size) || size == -1) {
    size = field.size;
    if (isNaN(size)) {
      var cols = field.cols;
      var rows = field.rows;
      if (!isNaN(cols) || !isNaN(rows)) {
        if (isNaN(cols)) {
          cols = 20;
        }
        if (isNaN(rows)) {
          rows = 2;
        }
        size = cols * rows;
      } else {
        size = 20;
      }
    }
  }
  nextFieldFocus( field, size, true, event );
}

function nextFocus( field, size, event )
{
  nextFieldFocus( field, size, true, event );
}

function nextFieldFocus( field, size, accents, event )
{
  /* Removendo acentuação caso seja indesejada */
  if( accents == false )
    field.value = removeAccents( field.value );

  /* Tenta recuperar a tecla pelo Netscape */
  var key = event.keyCode;
  /* ou pelo IE */
  if( key == null )
    key = event.which;
  /* Se não conseguir por nenhum dos dois, seta como A */
  if( key == null )
    key = 65;

  /* Se não for tecla de inserção de caracteres, retorna sem fazer o avanço */
  if ( (key == 9) /* TAB */
    || (key >= 16 && key <= 18) /* SHIFT, CTRL, ALT */
    || (key == 20) /* CAPSLCK */
    || (key >= 33 && key <= 40) /* PGUP, PGDWN, END, HOME, LEFT, UP, RIGHT, DOWN */
    || (key == 91 || key == 93) /* WINDOWS, POPUP */
    || (key == 144) /* NUMLCK */
    || (key >= 112 && key <= 123) ) /* F1 - F12 */
    return true;

  var i;
  var value = field.value;
  var selected = false; 

  if( field.form.elements.length != 0 &&
      size <= value.length &&
      key != 0 && key != 8 && key != 9 && key != 16 && key != 20 && key != 27 &&
      !(key >= 33 && key <= 46) &&
      !(key >= 16 && key <= 18) &&
      !(key >= 90 && key <= 93) &&
      !(key >= 112 && key <= 123) &&
      !(key >= 144 && key <= 145) )
    for( i = 0; i < field.form.elements.length - 1 && ! selected; i++ )
      if( field == field.form[ i ] )
        for( j = i + 1; j < field.form.elements.length && ! selected; j++ )
          if( field.form[ j ].type != "hidden" && field.form[ j ].disabled != true )
          {
            field.form[ j ].focus();
            selected = true
          }
  /* se não moveu o foco e já está no limite, retira o foco do componente */
  if (size <= value.length && !selected)
    field.blur();
}

function getField( c )
{
  var i;
  var j;
  for( i = 0; i < document.forms.length; i++ )
  {
    var f = document.forms[ i ];
    for( j = 0; j < f.elements.length; j++ )
    {
      var campo = f[ j ];
      if( c == campo.name )
        return campo;
    }
  }
  return null;
}

function setFocus( campofoco )
{
  var campo = getField( campofoco )
  if( campo != null && campo.disabled != true)
    campo.focus();
}

function setFirstFieldFocus()
{
  if( defaultFocusOccur != null )
    window.location.hash = defaultFocusOccur;
  if( defaultFocusField != null )
    setFocus( defaultFocusField );
  else
  {
    var form = document.forms[ 0 ];
    if( form != null )
      for( var i = 0; i < form.elements.length; ++i )
        if( form[ i ].type != 'hidden' && form[ i ].disabled != true)
        {
          form[ i ].focus();
          break;
        }
  }
}

function selectAll( newState )
{
  for( var i = 0; i < document.forms[0].elements.length; i++ )
    document.forms[0].elements[i].checked = newState;
}

function getQueryString()
{
  var res = "";
  for( var i = 0; i < document.forms[0].elements.length; i++ )
  {
    var field = document.forms[0].elements[i];
    if( i > 0 )
      res = res + "&";
    res = res + field.name + "=" + field.value;
  }
  return res;
}

function getQueryStringWithout( e )
{
  var res = "";
  for( var i = 0; i < document.forms[0].elements.length; i++ )
  {
    var elem = document.forms[0].elements[i];
    if( e != elem.name )
    {
      if( i > 0 )
        res = res + "&";
      res = res + elem.name + "=" + elem.value;
    }
  }
  return res;
}

function getQueryStringWithoutSubmits()
{
  var res = "";
  for( var i = 0; i < document.forms[0].elements.length; i++ )
  {
    var elem = document.forms[0].elements[i];
    if( elem.type != "submit" )
    {
      if( i > 0 )
        res = res + "&";
      res = res + elem.name + "=" + elem.value;
    }
  }
  return res;
}

function getWindowName()
{
  return window.name.length > 0 ? window.name : "_top";
}

function removeAccents( fieldValue )
{
  var value = '';
  var caracterStr;
  for( var j = 0; j < fieldValue.length; ++j )
  {
    caracterStr  = fieldValue.charAt( j );

    switch( caracterStr )
    {
      case 'ª':
      case 'º':
      case '\'':
      case '´': 
      case '`': 
      case '^': 
      case '¨': 
      case '~': break;

      case 'Ã':
      case 'Á':
      case 'À':
      case 'Â':
      case 'Ä':
        {
          value = value + 'A';
          break;
        }
      case 'ã':
      case 'á':
      case 'à':
      case 'â':
      case 'ä':
        {
          value = value + 'a';
          break;
        }
                
      case 'Ç':
        {
          value = value + 'C';
          break;
        }
      case 'ç':
        {
          value = value + 'c';
          break;
        }

      case 'É':
      case 'È':
      case 'Ê':
      case 'Ë':
        {
          value = value + 'E';
          break;
        }
      case 'é':
      case 'è':
      case 'ê':
      case 'ë':
        {
          value = value + 'e';
          break;
        }

      case 'Í':
      case 'Ì':
      case 'Î':
      case 'Ï':
        {
          value = value + 'I';
          break;
        }
      case 'í':
      case 'ì':
      case 'î':
      case 'ï':
        {
          value = value + 'i';
          break;
        }

      case 'Ý':
        {
          value = value + 'Y';
          break;
        }
      case 'ý':
        {
          value = value + 'y';
          break;
        }

      case 'Õ':
      case 'Ó':
      case 'Ò':
      case 'Ô':
      case 'Ö':
        {
          value = value + 'O';
          break;
        }
      case 'õ':
      case 'ó':
      case 'ò':
      case 'ô':
      case 'ö':
        {
          value = value + 'o';
          break;
        }

      case 'Ú':
      case 'Ù':
      case 'Û':
      case 'Ü':
        {
          value = value + 'U';
          break;
        }
      case 'ú':
      case 'ù':
      case 'û':
      case 'ü':
        {
          value = value + 'u';
          break;
        }

      case 'Ñ':
        {
          value = value + 'N';
          break;
        }
      case 'ñ':
        {
          value = value + 'n';
          break;
        }
      default: value = value + caracterStr;
    }
  }
  return( value );
}

/**
 * Format
 */

// Constantes
SYMBOL      = 0;
CARACTER    = 1;
SEPARATOR   = 2;
SIGNAL      = 3;
UPPER       = 4;
LOWER       = 5;
MINUS       = 6;
OTHER       = 7;


function fillLeft( str, c, len )
{
  for( var i = str.length; i < len; ++i )
    str = c + str;
  return( str );
}

function sizeMask( mask )
{
  var caracter;
  var lenMask = 0;
  var type;
  for( var i = 0; i < mask.length; ++i )
  {
    caracter = mask.charAt( i );
    type = findSymbol( caracter );
    if( type != UPPER &&
        type != LOWER &&
        type != SIGNAL &&
        type != MINUS &&
        type != SEPARATOR )
      ++lenMask;
  }
  return( lenMask );
}

function findSymbol( symbol )
{
  var typeSymbol = SYMBOL;
  switch( symbol )
  {
    case '#':
    case '0':
    case 'L':
    case 'l':
    case 'A':
    case 'a':
    case 'C':
    case 'c': {
                typeSymbol = CARACTER;
                break;
              }
    case 'S': {
                typeSymbol = SIGNAL;
                break;
              }
    case 's': {
                typeSymbol = MINUS;
                break;
              }
    case '>': { typeSymbol = UPPER;
                break;
              }
    case '<': { typeSymbol = LOWER;
                break;
              }
    case '\\': {
                 typeSymbol = SEPARATOR;
                 break;
               }
    default: typeSymbol = OTHER;
  }
  return( typeSymbol );
}

/**
 * Number
 */

function validateCaracterNumber( str )
{
  var caracterStr;
  for( var j = 0; j < str.length; ++j )
  {
    caracterStr  = str.charAt( j );
    var charCode = str.charCodeAt( j );
    //caracter numérico
    if( !isDigit( caracterStr ) &&
        caracterStr != '+' &&
        caracterStr != '-' )
      return( false );
  }//for
  return( true )
}

function numberZeros( displayMask )
{
  var number = 0;
  for( var i = 0; i < displayMask.length; ++i )
    if( displayMask.charAt( i ) == '0' )
      ++number;
  return( number );
}

function insertZeros( value, displayMask )
{
  var number = numberZeros( displayMask );
  if( value.length != 0 ) {
  	if ( number > 0 )
      value = fillLeft( deleteZerosLeft( value ), '0', number );
    else
      value = deleteZerosLeft( value );
  }
  return( value );
}

function deleteZerosLeft( value )
{
  var result = "";
  var i;
  for( i = 0; i < value.length; i++ )
  {
    var caracter = value.charAt( i );
    if( caracter != '0' )
    {
      if( isDigit( caracter ) )
        break;
      result = result + caracter;
    }
  }
  result = result + value.substring( i, value.length );
  if( result.length == 0 && value.indexOf( "0" ) != -1 )
    result = "0";
  return( result );
}

function deleteMaskNumber( value )
{
  var caracterValue;
  var valueDelete = "";
  for( var j = 0; j < value.length; ++j )
  {
    caracterValue = value.charAt( j );
    if( isDigit( caracterValue ) )
      valueDelete = valueDelete + caracterValue;
  }
  return( valueDelete );
}

function formatNumber( value, displayMask )
{
  //Verifica se o valor tem sinal e qual é o sinal
  var sinal      = "+";
  var sinalMinus = "";
//  if( value.indexOf( "-" ) != -1 )
  if( lastSignal( value ).indexOf( "-" ) > -1 )
  {
    sinal      = "-";
    sinalMinus = "-";
  }
  //Exclui símbolos do campo
  value = deleteMaskNumber( value );
  var formatValue = "";
  var caracter;
  var symbol;
  var anterior;
  //Coloca zeros no início da string
  value = insertZeros( value, displayMask );
  var lenValue = value.length - 1;
  //Percorre a máscara da direita para a esquerda
  for( var i = displayMask.length - 1; i > -1; --i )
  {
    caracter = displayMask.charAt( i );
    //Se já estiver após a primeira posição
    if ( i > 0 )
    {
      //Verifica se o anterior é um separador,
      //incluindo o caracter atual como símbolo do valor.
      anterior = findSymbol( displayMask.substring( i - 1, i ) );
      if ( anterior == SEPARATOR )
      {
        formatValue = caracter + formatValue;
        --i;
      }
    }
    if ( lenValue > -1)
    {
      symbol = findSymbol( caracter );
      if ( symbol == CARACTER )
      {
        formatValue = value.substring( lenValue, lenValue + 1 ) + formatValue;
        --lenValue;
      }
      else
      if ( symbol == SIGNAL )
        formatValue = sinal + formatValue;
	  else
        if ( symbol == MINUS )
          formatValue = sinalMinus + formatValue;
        else
          formatValue = caracter + formatValue;
    }
    else
      break;
  }
  if( hasMinusSignal( displayMask ) && formatValue.indexOf( "-" ) == -1 )
    formatValue = sinalMinus + formatValue;
  else
    if( hasSignal( displayMask ) && formatValue.indexOf( "+" ) == -1 && formatValue.indexOf( "-" ) == -1 )
      formatValue = sinal + formatValue;
  return( formatValue );
}

function lastSignal( value )
{
  var iPlus = value.lastIndexOf( "+" );
  var iMinus = value.lastIndexOf( "-" );
  if( ( iPlus == -1 && iMinus == -1 ) ||
      iPlus > iMinus )
    return "+";
  return "-";
}

function hasSignal( mask )
{
  return( mask.indexOf( "S" ) > -1 || mask.indexOf( "s" ) > -1 )
}

function hasMinusSignal( mask )
{
  return( mask.indexOf( "s" ) > -1 )
}

function formatValueNumber( field, displayMask, event )
{
  /* Tenta recuperar a tecla pelo Netscape */
  var key = event.keyCode;
  /* ou pelo IE */
  if( key == null )
    key = event.which;
  if( key != 9 )
  { 
    var value = field.value;
    var valueFormated = formatNumber( value, displayMask );
    if( value != valueFormated )
      field.value = valueFormated;
  }
}

/**
 * String
 */

function positionMaskString( offs, displayMask )
{
  var pos       = 0;
  var i         = 0;
  var tipo      = 1;
  var ultimoTipo;
  var caracter = " ";

  while ( i < displayMask.length )
  {
    //caracter da máscara
    caracter = displayMask.charAt( i );
    ultimoTipo = tipo;
    tipo = findSymbol( caracter );
    if ( pos >= offs &&
         tipo == CARACTER &&
         ultimoTipo != SEPARATOR )
      break;
    if ( tipo != SEPARATOR &&
         tipo != LOWER &&
         tipo != UPPER )
      ++pos;
    ++i;
  }
  return( caracter );
}

function validateCaracterString( str, offs, displayMask )
{
  var caracterMask;
  var caracterStr;
  var charCode;
  for( var j = 0; j < str.length; ++j )
  {
    caracterMask = positionMaskString( offs + j, displayMask );
    caracterStr  = str.charAt( j );
    charCode     = str.charCodeAt( j );
    switch( caracterMask )
    {
      case '#':
      case '9': { //caracter numérico ou espaço
                  if( !(isDigit( caracterStr ) || caracterStr == ' ') )
                    return( false );
                  break;
                }
      case '0': { //caracter numérico
                  if( !isDigit( caracterStr ) )
                    return( false );
                  break;
                }
      case 'A':
      case 'a': { //caracter alfanumérico
                  if( !(isLetterOrDigit( caracterStr ) || caracterStr == ' ') )
                    return( false );
                  break;
                }
      case 'L':
      case 'l': { //caracter alfabético
                  if( !(isLetter( caracterStr ) || caracterStr == ' ') )
                    return( false );
                  break;
                }
      case 'C':
      case 'c':  break;
      case 'S': {
                  if( caracterStr != '+' &&
                      caracterStr != '-' )
                    return( false );
                  break;
                }
      case '\\':
      default  : return( false );
    }
  }
  return( true );
}

function deleteMaskString( value, displayMask ) 
{
  var caracter      = "";
  var valueDelete   = "";
  var valor         = 0;
  var lenValue      = 0;
  var caracterValue = "";
  var i             = 0;

  while ( i < displayMask.length && lenValue < value.length )  
  {
    caracter      = displayMask.substr( i, 1 );
    caracterValue = value.substr( lenValue, 1 );
    valor         = findSymbol( caracter );
    if( valor == CARACTER || valor == SIGNAL )
      valueDelete = valueDelete + caracterValue;
    else if ( valor == UPPER || valor == LOWER )
      --lenValue;
    else if( valor == SEPARATOR )
    {
      if ( displayMask.charAt(i+1) != caracterValue )
        valueDelete = valueDelete + caracterValue;
      ++i;
    }
    else
      if ( caracter != caracterValue )
        valueDelete = valueDelete + caracterValue;
    ++lenValue;
    ++i;
  }
  return( valueDelete );
}

function formatString( value, displayMask )
{
  if ( value.length == 0 )
    return value;
  value = deleteMaskString( value, displayMask );
  var formatValue = "";
  var lenValue = 0;
  var caracter;
  var symbol;
  var i = 0;
  var typeCase = OTHER;

  while ( i < displayMask.length )
  {
    caracter = displayMask.charAt( i );

    symbol = findSymbol( caracter );
    if ( symbol == SEPARATOR )
    {
      formatValue =  formatValue + displayMask.substr( i + 1, 1 );
      ++i;
    }
    else
    if ( symbol == CARACTER )
    {
      if ( value.length == lenValue )
        break;
      var check = validateCaracterString( value.substr( lenValue, 1 ), formatValue.length, displayMask );

      if ( check )
      {
        var valueCaracter = value.substr( lenValue, 1 );
        if ( typeCase == UPPER )
          valueCaracter = valueCaracter.toUpperCase();
        else
          if ( typeCase == LOWER )
            valueCaracter = valueCaracter.toLowerCase()
        formatValue =  formatValue + valueCaracter;
      }
      else
        --i;
      ++lenValue;
    }
    else
    if ( symbol == UPPER )
      typeCase = UPPER;
    else
    {
      if ( symbol == LOWER )
      {
        if ( findSymbol( displayMask.substr( i + 1, 1 ) ) == UPPER )
        {
          typeCase = OTHER;
          ++i;
        }
        else
          typeCase = LOWER;
      }
      else
        if ( lenValue < value.length )
          formatValue =  formatValue + caracter;
    }
    ++i;
  }
  return( formatValue );
}

function formatValueString( field, displayMask, keyEvent )
{
  var value = field.value;
  var valueFormated = formatString( value, displayMask );
  if( value != valueFormated )
    field.value = valueFormated;
}

