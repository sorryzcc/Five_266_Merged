# Five_266_Merged
背景：可能存在很多文本当初复制了WW版的key，但在266又改过了文本内容，导致在MS合并后什么现在读取的是MS原本的旧的文本。当前只能发现一个改一个，效率低。

方案：
开发协助扫描ms和266，抽出 【key相同文本不同的条目】。po再各自确认文本，在文本管理平台新增正确文本，游戏内改key。

备注：扫描出来文本不一致的结果不可以直接覆盖MS。因为key本身海外也在用，改了文本内容会影响海外版，或者，出现海外版和大陆版同一条key但文本不同的情况。

python test.py 生成MS和266表的对比,并且取出负责人

