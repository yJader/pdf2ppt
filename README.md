# pdf2ppt

## 项目背景

使用LaTeX beamer制作了pdf格式的slides后, 如果需要演讲者视图功能, 在MacOS上[Présentation.app](http://iihm.imag.fr/blanch/software/osx-presentation/)是个不错的选择, 但是在Windows上的相似软件效果都不够好(且不够通用), 所以需要一个Windows上的解决方案.

这时候将它转为pptx格式, 使用PowerPoint/WPS的演讲者视图功能, 就可以在Windows上顺利的使用演讲者视图功能了

## 功能

将pdf转为pptx格式, 并将pdfnote转为pptx的备注

## 使用

```bash
# 查看更详细的帮助
python pdf2ppt.py -h 

# 转换  
python pdf2ppt.py convert -i input.pdf -o output.pptx --dpi 800

# 清理
python pdf2ppt.py clean 
```

## 备注

感谢gemini, 光速开发完成XD
