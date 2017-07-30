// dllmain.h : Declaration of module class.

class CTsComModule : public ATL::CAtlDllModuleT< CTsComModule >
{
public :
	DECLARE_LIBID(LIBID_TsComLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TSCOM, "{D93D61BB-0370-49B9-8133-2D642DD14AA0}")
};

extern class CTsComModule _AtlModule;
