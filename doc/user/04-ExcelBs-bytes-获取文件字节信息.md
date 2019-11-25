# 需求场景

有时候我们并不需要直接生成一个 excel 文件，可以根据 excel 内容，进行相关自定义操作。

比如生成一个 web 下载 excel 文件。

## 实现方式

可以返回一个文件流，但是个人感觉如果流忘记关闭等反而比较麻烦。

所以直接返回 excel 文件对应的文件内容 bytes[] 数组信息，用户可以根据该字节数组信息，进行相关操作。

## 支持版本

v0.0.6

# 测试用例

## 测试代码

```java
// 获取对应文件流
byte[] bytes = ExcelBs.newInstance()
        .append(User.buildUserList())
        .bytes();

// 根据文件内容，自行选择应用场景，到如 web 下载。
// 此处演示文件创建
final String filePath = PathUtil.getAppTestResourcesPath()+"/bytes.xls";
FileUtil.createFile(filePath, bytes);
```

## 简单说明

`ExcelBs.newInstance()` 这里直接创建一个引导类，因为不涉及到 excel 文件，可以暂时不指定文件路径。

通过 append() 方法写入对应的列表信息，然后就可以通过 bytes() 方法获取该 excel 对应的内容信息。

详细信息参见测试 ExcelBsTest#bytesTest()