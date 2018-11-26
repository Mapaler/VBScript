# 枫谷剑仙自编VBS脚本
打包下载项目请点右上角绿色的“Clone or download”按钮，然后选择“Download ZIP”
# 分项使用介绍
## 日文乱码系列
### 日文文本文件转换为UTF8编
将 Shift-JIS 编码的txt文件转换为 UTF-8 编码

下载 [日文乱码文本 to UTF8.vbs](https://github.com/Mapaler/VBScript/raw/master/%E6%97%A5%E6%96%87%E4%B9%B1%E7%A0%81%E6%96%87%E6%9C%AC%20to%20UTF8.vbs)  
将内容为乱码的txt文件（可多个）拖到本脚本上运行，会自动生成文件名后面加了“_U”的UTF-8编码文件。

*注：此脚本不改变文件扩展名。*
### 日文乱码文件名转换到Unicode改名
ZIP不支持Unicode字符，MMD素材等很多日本网友用zip打包的，用GBK编码解压后就是一堆乱码。如果没能使用BandZip等支持修改编码的压缩软件从源头上以Shift-JIS编码解压，也可以使用本工具对解压后的文件进行改名。

下载 [日文乱码文件名 to Unicode改名.bat](https://github.com/Mapaler/VBScript/raw/master/%E6%97%A5%E6%96%87%E4%B9%B1%E7%A0%81%E6%96%87%E4%BB%B6%E5%90%8D%20to%20Unicode%E6%94%B9%E5%90%8D.bat) 和 [VBScript/JP_fn_to_U.vbs](https://github.com/Mapaler/VBScript/raw/master/VBScript/JP_fn_to_U.vbs) 两个文件，注意要按照文件夹结构保存。  
将文件名为乱码的文件或文件夹（可多个）拖到本脚本上运行，会自动转换名称到原始的日文。  
如果不满意转换结果，可以进入`VBScript`文件夹使用对应生成的`恢复.bat`将修改的文件名还原。  
遇到错误将已是中文的文件名转换为乱码的情况，可对`JP_fn_to_U.vbs`进行修改，`NoCovertChar`内可以添加凡是遇到该字符就不对文件名进行转换的字。

*注：此脚本不改变文件内容。*
## 缩减文件夹
下了一堆无损/本子后，批量解压后有的文件夹内有多个层级，而且都是单个的文件夹，很没有意义。  
比如`(C86)[幽閉サテライト] 残響は鳴り止まず/(C86)[幽閉サテライト] 残響は鳴り止まず/残響は鳴り止まず`  
此工具将把仅有单个文件夹的多层缩减到一层`(C86)[幽閉サテライト] 残響は鳴り止まず`。

下载 [日文乱码文件名 to Unicode改名.bat](https://github.com/Mapaler/VBScript/raw/master/%E7%BC%A9%E5%87%8F%E6%96%87%E4%BB%B6%E5%A4%B9.bat) 和 [VBScript/ReduceFolder.vbs](https://github.com/Mapaler/VBScript/raw/master/VBScript/ReduceFolder.vbs) 两个文件，注意要按照文件夹结构保存。  
将所有文件夹所在的父文件夹拖到本脚本上运行。
## 固实7-Zip压缩包快速转JPG
下了一堆CG后，发现都是固始压缩的7z，解压比较缓慢，并且解压后都是bmp，图片很大。本程序可以自动将固实7z解压，并转换为jpg。

下载 [固实7-Zip压缩包快速转JPG.bat](https://github.com/Mapaler/VBScript/raw/master/%E5%9B%BA%E5%AE%9E7-Zip%E5%8E%8B%E7%BC%A9%E5%8C%85%E5%BF%AB%E9%80%9F%E8%BD%ACJPG.bat) 和 [VBScript/solid_7z_to_jpg_fast.vbs](https://github.com/Mapaler/VBScript/raw/master/VBScript/solid_7z_to_jpg_fast.vbs) 两个文件，注意要按照文件夹结构保存。  
将 [7z.exe](https://sparanoid.com/lab/7z/) 和 [nconvert.exe](https://www.xnview.com/en/nconvert/) 两个文件放到 VBScript 文件夹。  
使用[软媒内存盘](https://mofang.ruanmei.com/)建立一个内存盘可以加快速度（此项可以自行寻找其他替代软件），也可以用固态硬盘，使用机械硬盘也可以就是慢。编辑`solid_7z_to_jpg_fast.vbs`，修改`TempPath`到你的内存盘或者是其他缓存地点，修改`TempSpace`为缓存的可使用空间。  
将需要解压的7z压缩包拖到本脚本上运行。
## 文件夹显示伪名
使用Windows自带功能让英文的文件夹在资源管理器下显示为其他名字，而不修改其真实名称。

下载 [文件夹显示伪名.vbs](https://github.com/Mapaler/VBScript/raw/master/%E6%96%87%E4%BB%B6%E5%A4%B9%E6%98%BE%E7%A4%BA%E4%BC%AA%E5%90%8D.vbs)  
将文件夹拖上去
## 文字编码显示小工具系列
### 查询ASCLL码、查询Unicode码
查询字符的编码序号。
### 测试3DS可显示日文汉字
测试汉字在 Shift-JIS 编码内是否存在，存在的话在口袋妖怪XY里面就能显示。