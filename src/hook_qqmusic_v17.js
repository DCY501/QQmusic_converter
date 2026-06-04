const TARGET_DLL = "QQMusicCommon.dll";

var mod = Process.findModuleByName(TARGET_DLL);
if (!mod) {
    throw new Error("找不到 " + TARGET_DLL + "，请确保 QQ 音乐正在运行");
}

var EncAndDesMediaFileConstructorAddr = mod.findExportByName(
  "??0EncAndDesMediaFile@@QAE@XZ"
);

var EncAndDesMediaFileDestructorAddr = mod.findExportByName(
  "??1EncAndDesMediaFile@@QAE@XZ"
);

var EncAndDesMediaFileOpenAddr = mod.findExportByName(
  "?Open@EncAndDesMediaFile@@QAE_NPB_W_N1@Z"
);

var EncAndDesMediaFileGetSizeAddr = mod.findExportByName(
  "?GetSize@EncAndDesMediaFile@@QAEKXZ"
);

var EncAndDesMediaFileReadAddr = mod.findExportByName(
  "?Read@EncAndDesMediaFile@@QAEKPAEK_J@Z"
);

if (!EncAndDesMediaFileConstructorAddr || !EncAndDesMediaFileOpenAddr || !EncAndDesMediaFileReadAddr) {
    throw new Error("找不到 EncAndDesMediaFile 的关键函数，QQ 音乐版本可能不兼容");
}

var EncAndDesMediaFileConstructor = new NativeFunction(
  EncAndDesMediaFileConstructorAddr,
  "pointer",
  ["pointer"],
  "thiscall"
);

var EncAndDesMediaFileDestructor = new NativeFunction(
  EncAndDesMediaFileDestructorAddr,
  "void",
  ["pointer"],
  "thiscall"
);

var EncAndDesMediaFileOpen = new NativeFunction(
  EncAndDesMediaFileOpenAddr,
  "bool",
  ["pointer", "pointer", "bool", "bool"],
  "thiscall"
);

var EncAndDesMediaFileGetSize = new NativeFunction(
  EncAndDesMediaFileGetSizeAddr,
  "uint32",
  ["pointer"],
  "thiscall"
);

var EncAndDesMediaFileRead = new NativeFunction(
  EncAndDesMediaFileReadAddr,
  "uint",
  ["pointer", "pointer", "uint32", "uint64"],
  "thiscall"
);

rpc.exports = {
  decrypt: function (srcFileName, tmpFileName) {
    var EncAndDesMediaFileObject = Memory.alloc(0x28);
    EncAndDesMediaFileConstructor(EncAndDesMediaFileObject);

    var fileNameUtf16 = Memory.allocUtf16String(srcFileName);
    var opened = EncAndDesMediaFileOpen(EncAndDesMediaFileObject, fileNameUtf16, 1, 0);
    if (!opened) {
        EncAndDesMediaFileDestructor(EncAndDesMediaFileObject);
        throw new Error("Open 失败，文件可能不存在或无法解密");
    }

    var fileSize = EncAndDesMediaFileGetSize(EncAndDesMediaFileObject);
    if (fileSize === 0) {
        EncAndDesMediaFileDestructor(EncAndDesMediaFileObject);
        throw new Error("文件大小为 0");
    }

    var buffer = Memory.alloc(fileSize);
    var readSize = EncAndDesMediaFileRead(EncAndDesMediaFileObject, buffer, fileSize, 0);

    var data = buffer.readByteArray(fileSize);
    EncAndDesMediaFileDestructor(EncAndDesMediaFileObject);
    
    // 注意: File API 不支持中文路径，tmpFileName 必须是纯 ASCII
    var tmpFile = new File(tmpFileName, "wb");
    tmpFile.write(data);
    tmpFile.close();
    
    return { success: true, size: fileSize };
  },
};
