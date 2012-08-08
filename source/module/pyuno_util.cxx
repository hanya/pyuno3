/**************************************************************
 * 
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 * 
 *************************************************************/



#include "pyuno_impl.hxx"

#include <time.h>
#include <osl/thread.h>

#include <typelib/typedescription.hxx>

#include <rtl/strbuf.hxx>
#include <rtl/ustrbuf.hxx>
#include <osl/time.h>

#include <com/sun/star/beans/XMaterialHolder.hpp>

using rtl::OUStringToOString;
using rtl::OUString;
using rtl::OString;
using rtl::OStringBuffer;
using rtl::OUStringBuffer;


using com::sun::star::uno::TypeDescription;
using com::sun::star::uno::Sequence;
using com::sun::star::uno::Reference;
using com::sun::star::uno::XInterface;
using com::sun::star::uno::Any;
using com::sun::star::uno::Type;
using com::sun::star::uno::UNO_QUERY;
using com::sun::star::uno::TypeClass;
using com::sun::star::uno::RuntimeException;
using com::sun::star::uno::XComponentContext;
using com::sun::star::lang::XSingleServiceFactory;
using com::sun::star::script::XTypeConverter;
using com::sun::star::beans::XMaterialHolder;

#define USTR_ASCII(x) OUString( RTL_CONSTASCII_USTRINGPARAM( x ) )
namespace pyuno
{
PyRef ustring2PyUnicode( const OUString & str )
{
    PyRef ret;
#if PY_VERSION_HEX >= 0x03030000
    OString sUtf8(OUStringToOString(str, RTL_TEXTENCODING_UTF8));
    ret = PyRef( PyUnicode_DecodeUTF8( sUtf8.getStr(), sUtf8.getLength(), NULL) , SAL_NO_ACQUIRE );
#else
#if Py_UNICODE_SIZE == 2
    // YD force conversion since python/2 uses wchar_t
    ret = PyRef( PyUnicode_FromUnicode( (const Py_UNICODE*)str.getStr(), str.getLength() ), SAL_NO_ACQUIRE );
#else
    OString sUtf8(OUStringToOString(str, RTL_TEXTENCODING_UTF8));
    ret = PyRef( PyUnicode_DecodeUTF8( sUtf8.getStr(), sUtf8.getLength(), NULL) , SAL_NO_ACQUIRE );
#endif
#endif

    return ret;
}

#if PY_VERSION_HEX < 0x03000000
PyRef ustring2PyString( const OUString &str )
{
    OString o = OUStringToOString( str, osl_getThreadTextEncoding() );
    return PyRef( PyString_FromString( o.getStr() ), SAL_NO_ACQUIRE );
}
#endif

OUString pyString2ustring( PyObject *pystr )
{
    OUString ret;
    if( PyUnicode_Check( pystr ) )
    {
#if PY_VERSION_HEX >= 0x03030000
        Py_ssize_t size;
        char *pUtf8 = PyUnicode_AsUTF8AndSize(pystr, &size);
        ret = OUString(pUtf8, size, RTL_TEXTENCODING_UTF8);
#else
#if Py_UNICODE_SIZE == 2
	ret = OUString( (sal_Unicode * ) PyUnicode_AS_UNICODE( pystr ) );
#else
	PyObject* pUtf8 = PyUnicode_AsUTF8String(pystr);
	ret = OUString(PyString_AsString(pUtf8), PyString_Size(pUtf8), RTL_TEXTENCODING_UTF8);
	Py_DECREF(pUtf8);
#endif
#endif
    }
#if PY_VERSION_HEX < 0x03030000
    else
    {
        char *name = PyString_AsString(pystr );
        ret = OUString( name, strlen(name), osl_getThreadTextEncoding() );
    }
#endif
    return ret;
}

PyRef getObjectFromUnoModule( const Runtime &runtime, const char * func )
    throw ( RuntimeException )
{
    PyRef object(PyDict_GetItemString( runtime.getImpl()->cargo->getUnoModule().get(), (char*)func ) );
    if( !object.is() )
    {
        OUStringBuffer buf;
        buf.appendAscii( "couldn't find core function " );
        buf.appendAscii( func );
        throw RuntimeException(buf.makeStringAndClear(),Reference< XInterface >());
    }
    return object;
}


//------------------------------------------------------------------------------------
// Logging
//------------------------------------------------------------------------------------

bool isLog( RuntimeCargo * cargo, sal_Int32 loglevel )
{
    return cargo && cargo->logFile && loglevel <= cargo->logLevel;
}

void log( RuntimeCargo * cargo, sal_Int32 level, const rtl::OUString &logString )
{
    log( cargo, level, OUStringToOString( logString, osl_getThreadTextEncoding() ).getStr() );
}

void log( RuntimeCargo * cargo, sal_Int32 level, const char *str )
{
    if( isLog( cargo, level ) )
    {
        static const char *strLevel[] = { "NONE", "CALL", "ARGS" };

        TimeValue systemTime;
        TimeValue localTime;
        oslDateTime localDateTime;

        osl_getSystemTime( &systemTime );
        osl_getLocalTimeFromSystemTime( &systemTime, &localTime );
        osl_getDateTimeFromTimeValue( &localTime, &localDateTime );

        fprintf( cargo->logFile,
                 "%4i-%02i-%02i %02i:%02i:%02i,%03lu [%s,tid %ld]: %s\n",
                 localDateTime.Year,
                 localDateTime.Month,
                 localDateTime.Day,
                 localDateTime.Hours,
                 localDateTime.Minutes,
                 localDateTime.Seconds,
                 sal::static_int_cast< unsigned long >(
                     localDateTime.NanoSeconds/1000000),
                 strLevel[level],
                 sal::static_int_cast< long >(
                     (sal_Int32) osl_getThreadIdentifier( 0)),
                 str );
    }
}

namespace {

void appendPointer(rtl::OUStringBuffer & buffer, void * pointer) {
    buffer.append(
        sal::static_int_cast< sal_Int64 >(
            reinterpret_cast< sal_IntPtr >(pointer)),
        16);
}

}

void logException( RuntimeCargo *cargo, const char *intro,
                   void * ptr, const rtl::OUString &aFunctionName,
                   const void * data, const com::sun::star::uno::Type & type )
{
    if( isLog( cargo, LogLevel::CALL ) )
    {
        rtl::OUStringBuffer buf( 128 );
        buf.appendAscii( intro );
        appendPointer(buf, ptr);
        buf.appendAscii( RTL_CONSTASCII_STRINGPARAM("].") );
        buf.append( aFunctionName );
        buf.appendAscii( RTL_CONSTASCII_STRINGPARAM( " = " ) );
        buf.append(
            val2str( data, type.getTypeLibType(), VAL2STR_MODE_SHALLOW ) );
        log( cargo,LogLevel::CALL, buf.makeStringAndClear() );
    }

}

void logReply(
    RuntimeCargo *cargo,
    const char *intro,
    void * ptr,
    const rtl::OUString & aFunctionName,
    const Any &returnValue, 
    const Sequence< Any > & aParams )
{
    rtl::OUStringBuffer buf( 128 );
    buf.appendAscii( intro );
    appendPointer(buf, ptr);
    buf.appendAscii( RTL_CONSTASCII_STRINGPARAM("].") );
    buf.append( aFunctionName );
    buf.appendAscii( RTL_CONSTASCII_STRINGPARAM("()=") );
    if( isLog( cargo, LogLevel::ARGS ) )
    {
        buf.append(
            val2str( returnValue.getValue(), returnValue.getValueTypeRef(), VAL2STR_MODE_SHALLOW) );
        for( int i = 0; i < aParams.getLength() ; i ++ )
        {
            buf.appendAscii( RTL_CONSTASCII_STRINGPARAM(", " ) );
            buf.append(
                val2str( aParams[i].getValue(), aParams[i].getValueTypeRef(), VAL2STR_MODE_SHALLOW) );
        }
    }
    log( cargo,LogLevel::CALL, buf.makeStringAndClear() );
    
}

void logCall( RuntimeCargo *cargo, const char *intro,
              void * ptr, const rtl::OUString & aFunctionName,
              const Sequence< Any > & aParams )
{
    rtl::OUStringBuffer buf( 128 );
    buf.appendAscii( intro );
    appendPointer(buf, ptr);
    buf.appendAscii( RTL_CONSTASCII_STRINGPARAM("].") );
    buf.append( aFunctionName );
    buf.appendAscii( RTL_CONSTASCII_STRINGPARAM("(") );
    if( isLog( cargo, LogLevel::ARGS ) )
    {
        for( int i = 0; i < aParams.getLength() ; i ++ )
        {
            if( i > 0 )
                buf.appendAscii( RTL_CONSTASCII_STRINGPARAM(", " ) );
            buf.append(
                val2str( aParams[i].getValue(), aParams[i].getValueTypeRef(), VAL2STR_MODE_SHALLOW) );
        }
    }
    buf.appendAscii( RTL_CONSTASCII_STRINGPARAM(")") );
    log( cargo,LogLevel::CALL, buf.makeStringAndClear() );
}


}
