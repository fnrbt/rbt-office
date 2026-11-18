param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$OutputPath)
# The type library of an Office application as TSV, from the file that has it (Excel's is in EXCEL.EXE,
# Word's is MSWORD.OLB, both in the Office16 folder of a Click-to-Run Office):
#   member <TAB> interface <TAB> name <TAB> kind (method|get|let|set) <TAB> parameters <TAB> returns [<TAB> default]
#   coclass <TAB> name <TAB> default interface
#   enum <TAB> enum <TAB> name <TAB> value
#   alias <TAB> name <TAB> type
#   interface <TAB> name <TAB> extensible | nonextensible
# Parameters: name[?] (? = optional) separated by ";"; a type is the interface or the VARTYPE's name.
$ErrorActionPreference = 'Stop'
$path = $Path
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using ITypeLib = System.Runtime.InteropServices.ComTypes.ITypeLib;
using ITypeInfo = System.Runtime.InteropServices.ComTypes.ITypeInfo;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
public static class RbtTypeLib {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
    static extern int LoadTypeLibEx(string file, int kind, out ITypeLib lib);
    static string TypeName(ITypeInfo owner, TYPEDESC t) {
        var vt = (VarEnum)t.vt;
        if (vt == VarEnum.VT_PTR) return TypeName(owner, (TYPEDESC)Marshal.PtrToStructure(t.lpValue, typeof(TYPEDESC)));
        if (vt == VarEnum.VT_SAFEARRAY) return "Array";
        if (vt == VarEnum.VT_USERDEFINED) {
            ITypeInfo r; owner.GetRefTypeInfo((int)(long)t.lpValue, out r);
            string name, doc, help; int ctx; r.GetDocumentation(-1, out name, out doc, out ctx, out help);
            return name;
        }
        switch (vt) {
            case VarEnum.VT_I2: return "Integer"; case VarEnum.VT_I4: case VarEnum.VT_INT: return "Long";
            case VarEnum.VT_R4: return "Single"; case VarEnum.VT_R8: return "Double"; case VarEnum.VT_CY: return "Currency";
            case VarEnum.VT_DATE: return "Date"; case VarEnum.VT_BSTR: return "String"; case VarEnum.VT_DISPATCH: return "Object";
            case VarEnum.VT_BOOL: return "Boolean"; case VarEnum.VT_VARIANT: return "Variant"; case VarEnum.VT_UNKNOWN: return "Unknown";
            case VarEnum.VT_UI1: return "Byte"; case VarEnum.VT_VOID: case VarEnum.VT_HRESULT: return "";
            default: return vt.ToString();
        }
    }
    public static string Dump(string file) {
        ITypeLib lib; int hr = LoadTypeLibEx(file, 2, out lib); if (hr != 0) throw new Exception("LoadTypeLibEx " + hr);
        var sb = new StringBuilder();
        int count = lib.GetTypeInfoCount();
        for (int i = 0; i < count; i++) {
            ITypeInfo info; lib.GetTypeInfo(i, out info);
            string tname, doc, help; int ctx; info.GetDocumentation(-1, out tname, out doc, out ctx, out help);
            IntPtr pattr; info.GetTypeAttr(out pattr);
            var attr = (TYPEATTR)Marshal.PtrToStructure(pattr, typeof(TYPEATTR));
            try {
                if (attr.typekind == TYPEKIND.TKIND_ENUM) {
                    for (int v = 0; v < attr.cVars; v++) {
                        IntPtr pv; info.GetVarDesc(v, out pv);
                        var vd = (VARDESC)Marshal.PtrToStructure(pv, typeof(VARDESC));
                        string vname, vdoc, vhelp; int vctx; info.GetDocumentation(vd.memid, out vname, out vdoc, out vctx, out vhelp);
                        object value = Marshal.GetObjectForNativeVariant(vd.desc.lpvarValue);
                        sb.Append("enum\t" + tname + "\t" + vname + "\t" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) + "\n");
                        info.ReleaseVarDesc(pv);
                    }
                } else if (attr.typekind == TYPEKIND.TKIND_COCLASS) {
                    for (int k = 0; k < attr.cImplTypes; k++) {
                        IMPLTYPEFLAGS flags; info.GetImplTypeFlags(k, out flags);
                        if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) != 0 && (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) == 0) {
                            int href; info.GetRefTypeOfImplType(k, out href);
                            ITypeInfo r; info.GetRefTypeInfo(href, out r);
                            string rname, rdoc, rhelp; int rctx; r.GetDocumentation(-1, out rname, out rdoc, out rctx, out rhelp);
                            sb.Append("coclass\t" + tname + "\t" + rname + "\n");
                        }
                    }
                } else if (attr.typekind == TYPEKIND.TKIND_ALIAS) {
                    sb.Append("alias\t" + tname + "\t" + TypeName(info, attr.tdescAlias) + "\n");
                } else if (attr.typekind == TYPEKIND.TKIND_DISPATCH || attr.typekind == TYPEKIND.TKIND_INTERFACE) {
                    ITypeInfo target = info;
                    // A dual interface's dispatch view lists its members with their real parameters.
                    if (attr.typekind == TYPEKIND.TKIND_DISPATCH && (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FDUAL) != 0) {
                        try { int href; info.GetRefTypeOfImplType(-1, out href); info.GetRefTypeInfo(href, out target); } catch { target = info; }
                    }
                    // (VBA checks the member names of a non-extensible interface when it compiles; an extensible
                    // one - Excel's Range, Worksheet - may have members it does not list, found when they run.)
                    sb.Append("interface\t" + tname + "\t" + ((attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE) != 0 ? "nonextensible" : "extensible") + "\n");
                    IntPtr ptattr; target.GetTypeAttr(out ptattr);
                    var tattr = (TYPEATTR)Marshal.PtrToStructure(ptattr, typeof(TYPEATTR));
                    for (int f = 0; f < tattr.cFuncs; f++) {
                        IntPtr pf; target.GetFuncDesc(f, out pf);
                        var fd = (FUNCDESC)Marshal.PtrToStructure(pf, typeof(FUNCDESC));
                        if ((fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FRESTRICTED) == 0) {
                            var names = new string[fd.cParams + 1]; int got;
                            target.GetNames(fd.memid, names, names.Length, out got);
                            var ps = new List<string>();
                            string returns = TypeName(target, fd.elemdescFunc.tdesc);
                            int size = Marshal.SizeOf(typeof(ELEMDESC));
                            for (int p = 0; p < fd.cParams; p++) {
                                var ed = (ELEMDESC)Marshal.PtrToStructure(new IntPtr(fd.lprgelemdescParam.ToInt64() + p * size), typeof(ELEMDESC));
                                var pflags = ed.desc.paramdesc.wParamFlags;
                                bool optional = (pflags & PARAMFLAG.PARAMFLAG_FOPT) != 0 || (pflags & PARAMFLAG.PARAMFLAG_FHASDEFAULT) != 0;
                                bool retval = (pflags & PARAMFLAG.PARAMFLAG_FRETVAL) != 0;
                                bool lcid = (pflags & PARAMFLAG.PARAMFLAG_FLCID) != 0;
                                if (retval) { returns = TypeName(target, ed.tdesc); continue; }
                                if (lcid) continue;
                                string pname = (p + 1 < got && names[p + 1] != null) ? names[p + 1] : ("Arg" + (p + 1));
                                if (fd.invkind != INVOKEKIND.INVOKE_FUNC && fd.invkind != INVOKEKIND.INVOKE_PROPERTYGET && p == fd.cParams - 1) continue; // the value written
                                ps.Add(pname + (optional ? "?" : ""));
                            }
                            string kind = fd.invkind == INVOKEKIND.INVOKE_FUNC ? "method" : fd.invkind == INVOKEKIND.INVOKE_PROPERTYGET ? "get" : fd.invkind == INVOKEKIND.INVOKE_PROPERTYPUT ? "let" : "set";
                            // (DISPID 0: the default member - `coll(1)` is `coll.Item(1)`.)
                            sb.Append("member\t" + tname + "\t" + names[0] + "\t" + kind + "\t" + string.Join(";", ps) + "\t" + returns + (fd.memid == 0 ? "\tdefault" : "") + "\n");
                        }
                        target.ReleaseFuncDesc(pf);
                    }
                    target.ReleaseTypeAttr(ptattr);
                }
            } finally { info.ReleaseTypeAttr(pattr); }
        }
        return sb.ToString();
    }
}
'@
[System.IO.File]::WriteAllText($OutputPath, [RbtTypeLib]::Dump($path), (New-Object System.Text.UTF8Encoding($false)))
"$path -> $OutputPath"
