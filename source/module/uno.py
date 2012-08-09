#**************************************************************
#  
#  Licensed to the Apache Software Foundation (ASF) under one
#  or more contributor license agreements.  See the NOTICE file
#  distributed with this work for additional information
#  regarding copyright ownership.  The ASF licenses this file
#  to you under the Apache License, Version 2.0 (the
#  "License"); you may not use this file except in compliance
#  with the License.  You may obtain a copy of the License at
#  
#    http://www.apache.org/licenses/LICENSE-2.0
#  
#  Unless required by applicable law or agreed to in writing,
#  software distributed under the License is distributed on an
#  "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
#  KIND, either express or implied.  See the License for the
#  specific language governing permissions and limitations
#  under the License.
#  
#**************************************************************
import sys

import pyuno
import socket # since on Windows sal3.dll no longer calls WSAStartup
import importlib.abc
import types

# all functions and variables starting with a underscore (_) must be considered private
# and can be changed at any time. Don't use them
_g_ctx = pyuno.getComponentContext( )


def getComponentContext():
    """ returns the UNO component context, that was used to initialize the python runtime.
    """
    return _g_ctx

def getConstantByName( constant ):
    "Looks up the value of a idl constant by giving its explicit name"
    return pyuno.getConstantByName( constant )

def getTypeByName( typeName):
    """ returns a uno.Type instance of the type given by typeName. In case the
        type does not exist, a com.sun.star.uno.RuntimeException is raised.
    """
    return pyuno.getTypeByName( typeName )

def createUnoStruct( typeName, *args ):
    """creates a uno struct or exception given by typeName. The parameter args may
    1) be empty. In this case, you get a default constructed uno structure.
       ( e.g. createUnoStruct( "com.sun.star.uno.Exception" ) )
    2) be a sequence with exactly one element, that contains an instance of typeName.
       In this case, a copy constructed instance of typeName is returned
       ( e.g. createUnoStruct( "com.sun.star.uno.Exception" , e ) )
    3) be a sequence, where the length of the sequence must match the number of
       elements within typeName (e.g.
       createUnoStruct( "com.sun.star.uno.Exception", "foo error" , self) ). The
       elements with in the sequence must match the type of each struct element,
       otherwise an exception is thrown.
    """
    return getClass(typeName)( *args )

def getClass( typeName ):
    """returns the class of a concrete uno exception, struct or interface
    """
    return pyuno.getClass(typeName)

def isInterface( obj ):
    """returns true, when obj is a class of a uno interface"""
    return pyuno.isInterface( obj )

def generateUuid():
    "returns a 16 byte sequence containing a newly generated uuid or guid, see rtl/uuid.h "
    return pyuno.generateUuid()

def systemPathToFileUrl( systemPath ):
    "returns a file-url for the given system path"
    return pyuno.systemPathToFileUrl( systemPath )

def fileUrlToSystemPath( url ):
    "returns a system path (determined by the system, the python interpreter is running on)"
    return pyuno.fileUrlToSystemPath( url )

def absolutize( path, relativeUrl ):
    "returns an absolute file url from the given urls"
    return pyuno.absolutize( path, relativeUrl )

def getCurrentContext():
    """Returns the currently valid current context.
       see http://udk.openoffice.org/common/man/concept/uno_contexts.html#current_context
       for an explanation on the current context concept
    """
    return pyuno.getCurrentContext()

def setCurrentContext( newContext ):
    """Sets newContext as new uno current context. The newContext must
    implement the XCurrentContext interface. The implemenation should
    handle the desired properties and delegate unknown properties to the
    old context. Ensure to reset the old one when you leave your stack ...
    see http://udk.openoffice.org/common/man/concept/uno_contexts.html#current_context
    """
    return pyuno.setCurrentContext( newContext )


def hasModule(name):
    """ Check UNO module is there by its name. 
    
        Valid modules are module, constants and enum.
    """
    return pyuno.hasModule(name)


def getModuleElementNames(name):
    """ Get list of sub element names by name.
    
        Valid elements are module, interface, struct without template, 
        exception, enum and constants. And also list of names in enum and 
        constants can be taken.
    """
    return pyuno.getModuleElementNames(name)


class Enum:
    "Represents a UNO idl enum, use an instance of this class to explicitly pass a boolean to UNO"
    #typeName the name of the enum as a string
    #value    the actual value of this enum as a string
    def __init__(self,typeName, value):
        self.typeName = typeName
        self.value = value
        pyuno.checkEnum( self )

    def __repr__(self):
        return "<uno.Enum %s (%r)>" % (self.typeName, self.value)

    def __eq__(self, that):
        if not isinstance(that, Enum):
            return False
        return (self.typeName == that.typeName) and (self.value == that.value)

class Type:
    "Represents a UNO type, use an instance of this class to explicitly pass a boolean to UNO"
#    typeName                 # Name of the UNO type
#    typeClass                # python Enum of TypeClass,  see com/sun/star/uno/TypeClass.idl
    def __init__(self, typeName, typeClass):
        self.typeName = typeName
        self.typeClass = typeClass
        pyuno.checkType(self)
    def __repr__(self):
        return "<Type instance %s (%r)>" % (self.typeName, self.typeClass)

    def __eq__(self, that):
        if not isinstance(that, Type):
            return False
        return self.typeClass == that.typeClass and self.typeName == that.typeName

    def __hash__(self):
        return self.typeName.__hash__()


class Char:
    "Represents a UNO char, use an instance of this class to explicitly pass a char to UNO"
    # @param value pass a Unicode string with length 1
    def __init__(self,value):
        assert isinstance(value, str)
        assert len(value) == 1
        self.value=value

    def __repr__(self):
        return "<Char instance %s>" % (self.value, )

    def __eq__(self, that):
        if isinstance(that, str):
            if len(that) > 1:
                return False
            return self.value == that[0]
        if isinstance(that, Char):
            return self.value == that.value
        return False


class ByteSequence:
    def __init__(self, value):
        #if isinstance(value, str):
        if isinstance(value, bytes) or isinstance(value, bytearray):
            self.value = value
        elif isinstance(value, ByteSequence):
            self.value = value.value
        else:
            raise TypeError("expected string or bytesequence")

    def __repr__(self):
        return "<ByteSequence instance '%s'>" % (self.value, )

    def __eq__(self, that):
        if isinstance( that, ByteSequence):
            return self.value == that.value
        if isinstance(that, str):
            return self.value == that
        return False

    def __len__(self):
        return len(self.value)

    def __getitem__(self, index):
        return self.value[index]

    def __iter__( self ):
        return self.value.__iter__()

    def __add__( self , b ):
        if isinstance( b, str ):
            return ByteSequence( self.value + b )
        elif isinstance( b, ByteSequence ):
            return ByteSequence( self.value + b.value )
        raise TypeError( "expected string or ByteSequence as operand" )

    def __hash__( self ):
        return self.value.hash()


class Any:
    "use only in connection with uno.invoke() to pass an explicit typed any"
    def __init__(self, type, value ):
        if isinstance( type, Type ):
            self.type = type
        else:
            self.type = getTypeByName( type )
        self.value = value

def invoke( object, methodname, argTuple ):
    "use this function to pass exactly typed anys to the callee (using uno.Any)"
    return pyuno.invoke( object, methodname, argTuple )

#---------------------------------------------------------------------------------------
# don't use any functions beyond this point, private section, likely to change
#---------------------------------------------------------------------------------------

# private, referenced from the pyuno shared library
def _uno_struct__init__(self,*args):
    if len(args) == 1 and hasattr(args[0], "__class__") and args[0].__class__ == self.__class__ :
        self.__dict__["value"] = args[0]
    else:
        self.__dict__["value"] = pyuno._createUnoStructHelper(self.__class__.__pyunostruct__,args)

# private, referenced from the pyuno shared library
def _uno_struct__getattr__(self,name):
    return getattr(self.__dict__["value"],name)

# private, referenced from the pyuno shared library
def _uno_struct__setattr__(self,name,value):
    return setattr(self.__dict__["value"],name,value)

# private, referenced from the pyuno shared library
def _uno_struct__repr__(self):
    return repr(self.__dict__["value"])

def _uno_struct__str__(self):
    return str(self.__dict__["value"])

# private, referenced from the pyuno shared library
def _uno_struct__eq__(self,cmp):
    if hasattr(cmp,"value"):
        return self.__dict__["value"] == cmp.__dict__["value"]
    return False

def _uno_struct__dir__(self):
    return dir(self.__dict__["value"]) + list(self.__dict__.keys()) + \
                list(self.__class__.__dict__.keys())

# referenced from pyuno shared lib and pythonscript.py
def _uno_extract_printable_stacktrace( trace ):
    mod = None
    try:
        mod = __import__("traceback")
    except ImportError as e:
        pass
    ret = ""
    if mod:
        lst = mod.extract_tb( trace )
        max = len(lst)
        for j in range(max):
            i = lst[max-j-1]
            ret = ret + "  " + str(i[0]) + ":" + \
                  str(i[1]) + " in function " + \
                  str(i[2])  + "() [" + str(i[3]) + "]\n"
    else:
        ret = "Couldn't import traceback module"
    return ret


class UNOModule(types.ModuleType):
    """ Extended module class for UNO based modules. 
    
        Real value is not taken from pyuno until first request. 
        After first request of the value, it is kept as an attribute.
    """
    
    def __init__(self, fullname, loader):
        super().__init__(fullname)
        self.__file__ = "<" + fullname + ">"
        self.__loader__ = loader
        self.__path__ = fullname
        self.__package__ = ""
        self.__initializing__ = False
    
    def __getattr__(self, elt):
        value = None
        
        RuntimeException = pyuno.getClass("com.sun.star.uno.RuntimeException")
        try:
            value = pyuno.getClass(self.__path__ + "." + elt)
        except RuntimeException:
            try:
                value = Enum(self.__path__, elt)
            except RuntimeException:
                try:
                    value = pyuno.getConstantByName(self.__path__ + "." + elt)
                except RuntimeException:
                    if elt.startswith("typeOf"):
                        try:
                            value = pyuno.getTypeByName(self.__path__ + "." + elt[6:])
                        except RuntimeException:
                            raise AttributeError(
                                "type {}.{} is unknown".format(self.__path__, elt))
                    elif elt == "__all__":
                        try:
                            module_names = pyuno.getModuleElementNames(self.__path__)
                            self.__all__ = module_names
                            return module_names
                        except RuntimeException:
                            raise AttributeError("__all__")
                    else:
                        raise AttributeError(
                            "type {}.{} is unknown".format(self.__path__, elt))
        setattr(self, elt, value)
        return value


class UNOModuleLoader(importlib.abc.Loader):
    """ UNO module loader. 
    
        Creates new customized module for UNO if not yet loaded.
    """
    
    def load_module(self, fullname):
        mod = None
        if fullname in sys.modules:
            return sys.modules[fullname]
        else:
            mod = UNOModule(fullname, self)
        sys.modules.setdefault(fullname, mod)
        return mod


class UNOModuleFinder(importlib.abc.Finder):
    """ UNO module finder. 
    
        Generate module loader for UNO module. Valid module names are 
        one of module, enum and constants in IDL definitions.
    """
    
    LOADER = UNOModuleLoader()
    
    def find_module(self, fullname, path=None):
        if pyuno.hasModule(fullname):
            return self.__class__.LOADER
        return None


sys.meta_path.append(UNOModuleFinder())

