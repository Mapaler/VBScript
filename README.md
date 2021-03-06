# 枫谷剑仙自编VBS脚本
* 打包下载整个项目，请点右上角绿色的 **Code** 按钮，然后选择 [**Download ZIP**](https://github.com/Mapaler/VBScript/archive/master.zip)。  
![打包下载](document/download-zip.png)
* 选择单独下载某个文件时，在文件页面的 **Raw** 按钮上点击右键，选择另存为。  
![右键另存为](document/save-as.png)

# 分项使用介绍
## 一、日文乱码系列
### 1.日文文本文件转换为 UTF-8 编码
将 Shift-JIS 编码的txt文件转换为 UTF-8 编码

下载 [日文乱码文本 to UTF8.vbs](日文乱码文本%20to%20UTF8.vbs)  
将内容为乱码的txt文件（可多个）拖到本脚本上运行，会自动生成文件名后面加了“_U”的UTF-8编码文件。

*注：此脚本不改变文件扩展名。*
### 2.日文乱码文件名转换到 Unicode 改名
ZIP 不支持 Unicode 文件名，MMD 素材等很多日本网友用 zip 打包的，用 GBK 编码解压后就是一堆乱码。  
首先推荐使用[BandZip](http://www.bandisoft.com/bandizip/)等支持修改编码的压缩软件直接从源头上以 Shift-JIS 编码解压，也可以使用本工具对解压后的文件进行改名。  
本软件为中日文件名混合识别模式，可以自动区分文件名是否需要修改，但可能会有一定的识别错误。

下载 [日文乱码文件名 to Unicode改名.bat](日文乱码文件名%20to%20Unicode改名.bat) 和 [VBScript/JP_fn_to_U.vbs](VBScript/JP_fn_to_U.vbs) 两个文件，注意要按照文件夹结构保存。  
将文件名为乱码的文件或文件夹（可多个）拖到本脚本上运行，会自动转换名称到原始的日文。  
如果不满意转换结果，可以进入`VBScript`文件夹使用对应生成的`恢复.bat`将修改的文件名还原。  
遇到错误将已是中文的文件名转换为乱码的情况，可对`JP_fn_to_U.vbs`进行修改，`NoCovertChar`内可以添加凡是遇到该字符就不对文件名进行转换的字。

*注：此脚本不改变文件内容。*
## 二、缩减文件夹
下了一堆无损CD/本子后，批量解压后有的文件夹内有多个层级，而且都是单个的文件夹，很没有意义。  
比如`(C86)[幽閉サテライト] 残響は鳴り止まず/(C86)[幽閉サテライト] 残響は鳴り止まず/残響は鳴り止まず`  
此工具将把仅有单个文件夹的多层缩减到最外面一层`(C86)[幽閉サテライト] 残響は鳴り止まず`。

下载 [缩减文件夹.bat](缩减文件夹.bat) 和 [VBScript/ReduceFolder.vbs](VBScript/ReduceFolder.vbs) 两个文件，注意要按照文件夹结构保存。  
将所有文件夹所在的父文件夹拖到本脚本上运行。
## 三、固实7-Zip压缩包快速转JPG
下了一堆CG后，发现都是固始压缩的7z，解压比较缓慢，并且解压后都是bmp，图片很大。本程序可以自动将固实7z解压，并转换为jpg。

下载 [固实7-Zip压缩包快速转JPG.bat](固实7-Zip压缩包快速转JPG.bat) 和 [VBScript/solid_7z_to_jpg_fast.vbs](VBScript/solid_7z_to_jpg_fast.vbs) 两个文件，注意要按照文件夹结构保存。  
将 [7z.exe](https://sparanoid.com/lab/7z/) 和 [nconvert.exe](https://www.xnview.com/en/nconvert/) 两个文件放到 VBScript 文件夹。  
使用[软媒内存盘](https://mofang.ruanmei.com/)建立一个内存盘可以加快速度（此项可以自行寻找其他替代软件），也可以用固态硬盘，使用机械硬盘也可以就是慢。编辑`solid_7z_to_jpg_fast.vbs`，修改`TempPath`到你的内存盘或者是其他缓存地点，修改`TempSpace`为缓存的可使用空间。  
将需要解压的7z压缩包拖到本脚本上运行。
## 四、文件夹显示伪名
使用Windows自带功能让英文的文件夹在资源管理器下显示为其他名字，而不修改其真实名称。

下载 [文件夹显示伪名.vbs](文件夹显示伪名.vbs)  
将文件夹拖上去
## 五、文字编码显示小工具系列
### 1.查询ASCLL码、查询Unicode码
查询字符的编码序号。
### 2.测试3DS可显示日文汉字
测试汉字在 Shift-JIS 编码内是否存在，存在的话在口袋妖怪XY里面就能显示。

## 六、获取重定向链接
可以查询网址 30x 重定向的 URL


# 协议
本项目以GPLv3协议开源