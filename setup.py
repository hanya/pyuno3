
from distutils.core import setup, Extension

# check variables set by sdk
import sys, os, os.path

if not "OO_SDK_NAME" in os.environ:
    print("Please setup sdk and load its variables.")
    sys.exit(1)

base_path = os.environ["OFFICE_BASE_PROGRAM_PATH"]
ure_home = os.environ["OO_SDK_URE_HOME"]
sdk_home = os.environ["OO_SDK_HOME"]

sdk_inc = os.path.join(sdk_home, "include")

platform = sys.platform

if platform.startswith("linux"):
    # these variables can be found in make files of SDK.
    ure_types = os.path.join(ure_home, "share", "misc", "types.rdb")
    office_types = os.path.join(base_path, "offapi.rdb")
    
    sdk_lib = os.path.join(sdk_home, "lib")
    ure_lib = os.path.join(ure_home, "lib")
    stl_inc = os.path.join(sdk_inc, "stl")
    
    macros = [
        ("UNX", 1), 
        ("GCC", 1), 
        ("LINUX", 1), 
        ("CPPU_ENV", "gcc3"), 
        ("GXX_INCLUDE_PATH", os.environ["SDK_GXX_INCLUDE_PATH"])
    ]
    libs = [
        "uno_cppu", 
        "uno_cppuhelpergcc3", 
        "uno_sal", 
        "uno_salhelpergcc3", 
        "m", 
        "stlport_gcc", 
        "dl"
    ]
else:
    print("This platform is not supported yet. Please modify setup.py")
    sys.exit(1)



types = [
    "com.sun.star.uno.XComponentContext", 
    "com.sun.star.script.CannotConvertException", 
    "com.sun.star.lang.IllegalArgumentException", 
    "com.sun.star.lang.XServiceInfo", 
    "com.sun.star.lang.XTypeProvider", 
    "com.sun.star.beans.XPropertySet", 
    "com.sun.star.beans.XMaterialHolder", 
    "com.sun.star.beans.MethodConcept", 
    "com.sun.star.reflection.XIdlReflection", 
    "com.sun.star.reflection.XIdlClass", 
    "com.sun.star.reflection.XTypeDescription", 
    "com.sun.star.reflection.XEnumTypeDescription", 
    "com.sun.star.reflection.XConstantsTypeDescription", 
    "com.sun.star.reflection.XModuleTypeDescription", 
    "com.sun.star.reflection.XStructTypeDescription", 
    "com.sun.star.registry.InvalidRegistryException", 
    "com.sun.star.beans.XIntrospection", 
    "com.sun.star.script.XTypeConverter", 
    "com.sun.star.script.XInvocation2", 
    "com.sun.star.script.XInvocationAdapterFactory2", 
    "com.sun.star.container.XHierarchicalNameAccess", 
    "com.sun.star.lang.XUnoTunnel", 
    "com.sun.star.lang.XSingleServiceFactory", 
    "com.sun.star.uno.XWeak", 
    "com.sun.star.uno.XAggregation", 
    "com.sun.star.lang.XMultiServiceFactory", 
    "com.sun.star.uno.XCurrentContext", 
    
]

build_dir = os.path.join(".", "build")
types_inc = os.path.join(".", "build", "inc")

cppumaker_flag = os.path.join(".", "build", "flag")
if not os.path.exists(cppumaker_flag):
    
    if not os.path.exists(build_dir):
        os.mkdir(build_dir)
    
    import subprocess
    subprocess.Popen(
        [
            "cppumaker", 
            "-G", 
            "-BUCR", 
            "-O%s" % types_inc, 
            "-T%s" % ";".join(types), 
            ure_types, 
            office_types, 
        ]
    )
    
    open(cppumaker_flag, "w").close()



source_dir = "./source/module"
files = (
    "pyuno.cxx", 
    "pyuno_adapter.cxx", 
    "pyuno_callable.cxx", 
    "pyuno_except.cxx", 
    "pyuno_gc.cxx", 
    "pyuno_module.cxx", 
    "pyuno_runtime.cxx", 
    "pyuno_type.cxx", 
    "pyuno_util.cxx", 
)
cpp_files = [source_dir + "/" + i for i in files]


setup(
    name="pyuno", 
    version="3.4.0", 
    description="Python-UNO binding", 
    #package_dir={source_dir.replace("/", "."): ""}, 
    #py_modules=["uno", "unohelper"], 
    ext_modules=[
        Extension(
            "pyuno", 
            cpp_files, 
            define_macros=macros, 
            include_dirs=[
                "./inc", 
                types_inc, 
                sdk_inc, 
                source_dir, 
                stl_inc
            ], 
            library_dirs=[sdk_lib, ure_lib], 
            libraries=libs
        )
    ]
)
