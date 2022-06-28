# TyporaLuaFilter

- 对Typora的右键菜单添加颜色
- 支持导出带颜色的word文字和指定word中表格的边框样式



#### 对Typora的右键添加颜色

1. 下载myMenu.js

2. 找到Typora安装目录，找到resources/window.html，在文件的最后，</body>前添加：

3. ```html
    <script type="text/javascript" src="your file path/myMenu.js"></script>
    ```

4. 重启Typora，右键查看效果：

![](https://raw.githubusercontent.com/ZhouJunjun/image/master/markdown/lQLPJxZyFy1_-9DNAVfM_LBCau_LxRQqcAK7_gMIQCcA_252_343.png)

![image-20220628165423705](https://raw.githubusercontent.com/ZhouJunjun/image/master/markdown/image-20220628165423705.png)

5. 可以修改myMenu.js中的myColors添加/删除喜欢的颜色

#### 导出带颜色的word文字和指定word中表格的边框样式

1. 下载filter.lua，并配置导出参数：

    ```
    --lua-filter your file path/filter.lua
    ```

    

2. 正常导出word。可以下载test.md测试一下。对应的word效果：

    ![image-20220628171525568](https://raw.githubusercontent.com/ZhouJunjun/image/master/markdown/image-20220628171525568.png)