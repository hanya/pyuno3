#-*- encoding: utf-8 -*-

import sys
import os.path
import platform

major = sys.version_info[0]
minor = sys.version_info[1]
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.append(os.path.join(path, "source/module"))
sys.path.append(
    os.path.join(path, 
        "build/lib.{}-{}-{}.{}".format(
            platform.system().lower(), 
            platform.machine(), 
            major, minor)))

try:
    import uno
except Exception as e:
    import traceback
    traceback.print_exc()
    sys.exit(1)


def connect(uno_url):
    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        return resolver.resolve(uno_url)
    except Exception as e:
        print("Error on connect to office: " + uno_url)
        print(e)
        raise


import unittest


class PyUNOTestFunctions(unittest.TestCase):
    
    UNO_URL = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    
    ctx = None
    doc = None
    
    def setUp(self):
        if not self.__class__.ctx:
            self.__class__.ctx = connect(self.UNO_URL)
            self.__class__.doc = self.create_doc()
            #self.assertIsNotNone(self.ctx)
    
    def get_ctx(self):
        return self.__class__.ctx
    
    def get_doc(self):
        return self.__class__.doc
    
    def get_desktop(self):
        return self.create("com.sun.star.frame.Desktop")
    
    def create(self, name, args=None):
        """ Create instance of service specified by name. """
        smgr = self.get_ctx().getServiceManager()
        if args:
            return smgr.createInstanceWithArgumentsAndContext(name, args, self.get_ctx())
        else:
            return smgr.createInstanceWithContext(name, self.get_ctx())
    
    def create_doc(self):
        return self.get_desktop().loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, ())
    
    def test_getServiceManager(self):
        smgr = self.get_ctx().getServiceManager()
        self.assertIsNotNone(smgr)
    
    # functions defined in uno module
    
    def test_getComponentContext(self):
        r = uno.getComponentContext()
        self.assertIsNotNone(r)
    
    def test_getConstantByName(self):
        n = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        self.assertEqual(n, 150.0)
        n = uno.getConstantByName("com.sun.star.awt.FontSlant.ITALIC")
        self.assertEqual(n, 2)
        #uno.getConstantByName("ふぉお") # error ok
    
    def test_getTypeByName(self):
        
        def test(type_name, type_class):
            t = uno.getTypeByName(type_name)
            self.assertTrue(isinstance(t, uno.Type))
            self.assertEqual(t.typeName, type_name)
            self.assertTrue(isinstance(t.typeClass, uno.Enum))
            self.assertEqual(t.typeClass.value, type_class)
        
        test("long", "LONG")
        test("[]long", "SEQUENCE")
        test("com.sun.star.awt.XMouseListener", "INTERFACE")
    
    def test_createUnoStruct(self):
        from com.sun.star.awt import Rectangle
        rect1 = uno.createUnoStruct("com.sun.star.awt.Rectangle")
        self.assertTrue(isinstance(rect1, Rectangle))
        rect2 = uno.createUnoStruct("com.sun.star.awt.Rectangle", 100, 200, 50, 1)
        self.assertEqual(rect2.X, 100)
        rect3 = uno.createUnoStruct("com.sun.star.awt.Rectangle", rect2)
        #self.assertEqual(rect2, rect3)
    
    def test_getClass(self):
        from com.sun.star.uno import Exception as UNOException
        ex = uno.getClass("com.sun.star.uno.Exception")
        self.assertEqual(ex, UNOException)
        
        
    def test_isInterface(self):
        from com.sun.star.awt import XMouseListener
        from com.sun.star.awt.FontSlant import ITALIC
        
        self.assertTrue(uno.isInterface(XMouseListener))
        self.assertFalse(uno.isInterface(ITALIC))
    
    def test_generateUuid(self):
        v = uno.generateUuid()
        self.assertTrue(isinstance(v, uno.ByteSequence))
        self.assertEqual(len(v), 16)
        self.assertTrue(isinstance(v.value, bytes))
        
    
    def test_systemPathToFileUrl(self):
        sys_path = "/home/foo/bar/hoge/123.ods"
        url = "file:///home/foo/bar/hoge/123.ods"
        
        result = uno.systemPathToFileUrl(sys_path)
        self.assertEqual(result, url)
        
    
    def test_fileUrlToSystemPath(self):
        sys_path = "/home/foo/bar/hoge/123.ods"
        url = "file:///home/foo/bar/hoge/123.ods"
        
        result = uno.fileUrlToSystemPath(url)
        self.assertEqual(result, sys_path)
    
    def test_absolutize(self):
        relative_path = "../.."
        url = "file:///home/foo/bar/hoge/123.ods"
        desired = "file:///home/foo/bar/"
        result = uno.absolutize(url, relative_path)
        self.assertEqual(result, desired)
    
    def test_hasModule(self):
        import pyuno
        self.assertTrue(pyuno.hasModule("com"))
        self.assertTrue(pyuno.hasModule("com.sun"))
        self.assertTrue(pyuno.hasModule("com.sun.star.awt.FontWeight"))
        self.assertTrue(pyuno.hasModule("com.sun.star.awt.FontSlant"))
        self.assertFalse(pyuno.hasModule("foo"))
    
    def test_getModuleElementNames(self):
        self.assertTrue("sun" in uno.getModuleElementNames("com"))
        self.assertTrue("beans" in uno.getModuleElementNames("com.sun.star"))
        _all = set(uno.getModuleElementNames("com.sun.star.beans"))
        self.assertTrue("XExactName" in _all)
        self.assertTrue("NamedValue" in _all)
        self.assertTrue("UnknownPropertyException" in _all)
        self.assertTrue("PropertyState" in _all)
        self.assertTrue("PropertyAttribute" in _all)
        
        self.assertFalse("Introspection" in _all)
        self.assertFalse("Optional" in _all)
        self.assertFalse("PropertyValues" in _all)
        
        _all = set(uno.getModuleElementNames("com.sun.star.awt.FontSlant"))
        self.assertTrue("ITALIC" in _all)
        _all = set(uno.getModuleElementNames("com.sun.star.awt.FontWeight"))
        self.assertTrue("BOLD" in _all)
    
    # classes defined in uno module
    
    def test_Enum(self):
        repr_base = "<uno.Enum {type_name} ('{value}')>"
        type_name = "com.sun.star.awt.FontSlant"
        value = "ITALIC"
        repr_desired = repr_base.format(type_name=type_name, value=value)
        
        e = uno.Enum(type_name, value)
        self.assertEqual(repr(e), repr_desired)
        e2 = uno.Enum(type_name, value)
        self.assertTrue(e == e2)
        
        type_name2 = "com.sun.star.beans.PropertyState"
        value2 = "DIRECT_VALUE"
        repr_desired2 = repr_base.format(type_name=type_name2, value=value2)
        em = uno.Enum(type_name2, value2)
        self.assertEqual(repr(em), repr_desired2)
        
        self.assertFalse(e == em)
        self.assertTrue(e != em)
        
        # ToDo illegal type name and value
        
    
    def test_Type(self):
        repr_base = "<Type instance {type_name} ({type_class})>"
        
        type_name = "boolean"
        type_class = uno.Enum("com.sun.star.uno.TypeClass", "BOOLEAN")
        repr_desired = repr_base.format(type_name=type_name, type_class=type_class)
        t = uno.Type(type_name, type_class)
        self.assertTrue(t == uno.getTypeByName(type_name))
        self.assertFalse(t == uno.getTypeByName("void"))
        self.assertEqual(repr(t), repr_desired)
        self.assertEqual(hash(t), hash(type_name))
    
    def test_Char(self):
        repr_base = "<Char instance {}>"
        v = "c"
        repr_desired = repr_base.format(v)
        
        c = uno.Char(v)
        self.assertEqual(repr(c), repr_desired)
        self.assertEqual(c, v)
        self.assertEqual(c, c)
        self.assertNotEqual(c, uno.Char("v"))
    
    def test_ByteSequence(self):
        a = b"abcdef"
        b = b"xyz"
        
        bsa = uno.ByteSequence(a)
        bsb = uno.ByteSequence(b)
        self.assertEqual(bsa.value, a)
        self.assertEqual(bsb.value, b)
        self.assertEqual(bsa, a)
        self.assertEqual(bsa, bytearray(a))
        self.assertEqual(len(bsa), len(a))
        self.assertEqual(bsa[1], a[1])
        c = a + b
        bsc = uno.ByteSequence(c)
        self.assertEqual(bsc, bsa + bsb)
        self.assertEqual(bsc, c)
    
    def test_Any(self):
        vt = self.create_value_test()
        uno.invoke(vt, "setLong", (uno.Any("long", 100),))
        ret = uno.invoke(vt, "getLong", ())
        self.assertEqual(ret, 100)
    
    # import test
    def test_import_interface(self):
        from com.sun.star.container import XNameAccess, XIndexAccess
        self.assertTrue(uno.isInterface(XNameAccess))
        self.assertTrue(uno.isInterface(XIndexAccess))
    
    def test_import_struct(self):
        from com.sun.star.awt import Rectangle
        r = Rectangle()
        self.assertTrue(isinstance(r, Rectangle))
        self.assertTrue(isinstance(r, uno.UNOStruct))
        self.assertEqual(Rectangle.typeName, "com.sun.star.awt.Rectangle")
        self.assertEqual(Rectangle.__pyunostruct__, "com.sun.star.awt.Rectangle")
    
    def test_import_exception(self):
        from com.sun.star.uno import RuntimeException
        e = RuntimeException()
        self.assertTrue(isinstance(e, RuntimeException))
        self.assertTrue(isinstance(e, uno.UNOException))
        self.assertTrue(isinstance(e, Exception))
        self.assertEqual(RuntimeException.typeName, "com.sun.star.uno.RuntimeException")
        self.assertEqual(RuntimeException.__pyunostruct__, "com.sun.star.uno.RuntimeException")
    
    def test_import_enum(self):
        from com.sun.star.awt.FontSlant import OBLIQUE
        e = uno.Enum("com.sun.star.awt.FontSlant", "OBLIQUE")
        self.assertEqual(OBLIQUE, e)
    
    def test_import_constant(self):
        from com.sun.star.awt.FontWeight import BLACK
        self.assertEqual(BLACK, 200.0)
        import com.sun.star.awt.PosSize as PosSize
        self.assertEqual(PosSize.X, 1)
    
    def test_import_typeOf(self):
        from com.sun.star.container import typeOfXNameAccess
        t = uno.getTypeByName("com.sun.star.container.XNameAccess")
        self.assertEqual(typeOfXNameAccess, t)
    
    def test_import_module(self):
        import com.sun.star
        self.assertTrue(isinstance(com, uno.UNOModule))
        self.assertTrue(isinstance(com.sun, uno.UNOModule))
        self.assertTrue(isinstance(com.sun.star, uno.UNOModule))
    
    def test_import_unknown_module(self):
        def _import():
            import com.foo
        
        self.assertRaises(ImportError, _import)
    
    def test_import_unknown_atrribute(self):
        def _import():
            from com.sun.star.awt.FontSlant import FOO
        
        self.assertRaises(ImportError, _import)
    
    def test_import_imported_element(self):
        from com.sun.star.awt import FontSlant
        self.assertEqual(FontSlant.ITALIC, 
            uno.Enum("com.sun.star.awt.FontSlant", "ITALIC"))
        self.assertEqual(FontSlant.OBLIQUE, 
            uno.Enum("com.sun.star.awt.FontSlant", "OBLIQUE"))
        
        from com.sun.star.awt import XActionListener
        import com.sun.star.awt
        self.assertTrue(hasattr(com.sun.star.awt, "XActionListener"))
        # this is valid because hasattr calls __getattr__
        self.assertTrue(hasattr(com.sun.star.awt, "XButton"))
    
    #def test_import_all(self):
    #    pass
    
    # type class test
    
    def create_value_test(self, **kwds):
        from mytools import Values
        vs = Values()
        for k, v in kwds.items():
            setattr(vs, k, v)
        vt = self.create("mytools.ValueTest", (vs,))
        return vt
    
    def test_void(self):
        vt = self.create_value_test()
        self.assertIsNone(vt.getVoid())
    
    def test_char(self):
        vt = self.create_value_test(CharValue=uno.Char("c"))
        self.assertEqual(vt.getChar(), uno.Char("c"))
        vt.setChar(uno.Char("b"))
        self.assertEqual(vt.getChar(), uno.Char("b"))
        s = "あ"
        vt.setChar(s)
        self.assertEqual(vt.getChar(), uno.Char(s))
    
    def test_boolean(self):
        vt = self.create_value_test(BooleanValue=False)
        self.assertFalse(vt.getBoolean())
        vt.setBoolean(True)
        self.assertTrue(vt.getBoolean())
        vt.setBoolean(False)
        self.assertFalse(vt.getBoolean())
    
    def test_byte(self):
        vt = self.create_value_test(ByteValue=100)
        self.assertEqual(vt.getByte(), 100)
        vt.setByte(10)
        self.assertEqual(vt.getByte(), 10)
    
    def test_short(self):
        vt = self.create_value_test(ShortValue=30000)
        self.assertEqual(vt.getShort(), 30000)
        vt.setShort(-10000)
        self.assertEqual(vt.getShort(), -10000)
    
    def test_long(self):
        vt = self.create_value_test(LongValue=1234567)
        self.assertEqual(vt.getLong(), 1234567)
        vt.setLong(-1234567)
        self.assertEqual(vt.getLong(), -1234567)
    
    def test_hyper(self):
        v = 1111222333
        vt = self.create_value_test(HyperValue=v)
        self.assertEqual(vt.getHyper(), v)
        vt.setHyper(-v)
        self.assertEqual(vt.getHyper(), -v)
    
    #def test_float(self):
        #vt = self.create_value_test(FloatValue=100.111)
        #self.assertEqual(vt.getFloat(), 100.111)
        #vt.setFloat(-100.011)
        #self.assertEqual(vt.getFloat(), -100.011)
        # error on conversion between double and float?
    
    def test_double(self):
        vt = self.create_value_test(DoubleValue=100.111)
        self.assertEqual(vt.getDouble(), 100.111)
        vt.setDouble(-100.011)
        self.assertEqual(vt.getDouble(), -100.011)
    
    def test_string(self):
        vt = self.create_value_test(StringValue="hoge")
        self.assertEqual(vt.getString(), "hoge")
        s = "ｐｙてょｎ"
        vt.setString(s)
        self.assertEqual(vt.getString(), s)
        
        s = "マルチバイトテキスト Multi-byte text"
        doc = self.get_doc()
        text = doc.getText()
        text.setString(s)
        self.assertEqual(text.getString(), s)
    
    def test_type(self):
        vt = self.create_value_test(TypeValue=uno.getTypeByName("[]long"))
        self.assertEqual(vt.getType(), uno.getTypeByName("[]long"))
        vt.setType(uno.getTypeByName("com.sun.star.awt.Rectangle"))
        self.assertEqual(vt.getType(), uno.getTypeByName("com.sun.star.awt.Rectangle"))
    
    #def test_any(self):
    #    pass
    
    def test_enum(self):
        italic = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")
        doc = self.get_doc()
        text = doc.getText()
        text.setString("From PyUNO")
        cursor = text.createTextCursor()
        cursor.goRight(5, True)
        cursor.CharPosture = italic
        self.assertEqual(cursor.CharPosture, italic)
    
    def test_struct(self):
        from com.sun.star.table import BorderLine
        border = BorderLine(0xff0000, 10, 0, 0)
        doc = self.get_doc()
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.BottomBorder = border
        border2 = cursor.BottomBorder
        self.assertEqual(border.Color, border2.Color)
    
    def test_exception(self):
        pass
    
    def test_sequence(self):
        doc = self.get_doc()
        text = doc.getText()
        table = doc.createInstance("com.sun.star.text.TextTable")
        table.setName("NewTable")
        table.initialize(2, 2)
        text.insertTextContent(text.getEnd(), table, True)
        a = (("foo", "bar"), (1, 2))
        table.setDataArray(a)
        data = table.getDataArray()
        self.assertEqual(data, a)
        
        
        from com.sun.star.beans import PropertyValue
        arg1 = PropertyValue()
        arg1.Name = "InputStream"
        arg2 = PropertyValue()
        arg2.Name = "FilterName"
        arg2.Value = "writer_web_HTML"
        bs = b"<html><body><p>Text from <b>HTML</b>.</p></body></html>"
        
        sequence = self.create("com.sun.star.io.SequenceInputStream")
        sequence.initialize((uno.ByteSequence(bs),))
        arg1.Value = sequence
        
        text = doc.getText()
        text.setString("")
        
        text.getEnd().insertDocumentFromURL("", (arg1, arg2))
        sequence.closeInput()
    
    def test_stream(self):
        path = "/home/asuka/foo.txt"
        b = b"test text"
        
        pipe = self.create("com.sun.star.io.Pipe")
        pipe.writeBytes(uno.ByteSequence(b))
        pipe.flush()
        pipe.closeOutput()
        n, d = pipe.readBytes(None, 100)
        v = d.value
        pipe.closeInput()
        self.assertEqual(v, b)
    
    def test_interface(self):
        n = 0
        doc = self.get_doc()
        self.assertTrue(len(dir(doc)))
        desktop = self.get_desktop()
        frames = desktop.getFrames()
        for i in range(frames.getCount()):
            frame = frames.getByIndex(i)
            model = frame.getController().getModel()
            if doc == model:
                n += 1
        self.assertEqual(n, 1)
        self.assertFalse(doc == desktop)
    
    def test_dialog(self):
        return # needs dialog and user interaction
        from com.sun.star.awt import XActionListener
        import unohelper
        class ActionListener(unohelper.Base, XActionListener):
            def disposing(self, ev): pass
            def actionPerformed(self, ev):
                d.endExecute()
        
        dp = self.create("com.sun.star.awt.DialogProvider")
        d = dp.createDialog("vnd.sun.star.script:Standard.Dialog1?location=application")
        l = ActionListener()
        d.getControl("CommandButton1").addActionListener(l)
        d.execute()
        d.dispose()

    
if __name__ == "__main__":
    if sys.version_info[0] == 3:
        unittest.main()
    else:
        print("This unittest can execute only on Python3.")
        sys.exit(1)

