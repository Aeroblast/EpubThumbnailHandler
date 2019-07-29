# EpubThumbnailHandler
 Windows Shell Extention Thumbnail Handler for .epub


 ## 编译
 打开Visual Studio点编译。调试也行不过会报不能执行（当然不能直接执行）。

 ## 使用

 以下两个指令分别是安装和卸载，需要管理员权限执行。一旦安装，dll文件将被系统占用，无法修改，必须执行卸载才可修改。卸载后可能需要等一段时间系统才会取消占用，不想等可以重启。
``` Regsvr32.exe  EpubShellExtThumbnailHandler.dll ```
``` Regsvr32.exe /u EpubShellExtThumbnailHandler.dll ```

 ## 说明
 魔改自官方教程。
 谁知道会不会有什么内存泄漏、加载变慢之类的问题呢……所以暂时不提供编译版，至少会编译才能处理突发问题吧（
 测试环境Windows 10 1803，编译VS2017。

