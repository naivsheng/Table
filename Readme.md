# 思路流程

## 概论

files = "KW{woche} Bestellung KW{woche+1} Lieferung Übersicht.xlsx"

允许使用者直接更改excel表，删减供货商、门店。获取总览表中的行列信息，更新供货商、门店列表

减少不必要的操作，精简源代码

## 操作逻辑

将本次操作信息写入总览表，提示成功；若未能成功写入则弹窗提示，询问是否将数据暂存

数据暂存的格式为dist：{当前选定的供应商/门店:[本次操作的列表,操作标签项,要写入的标记]}

允许用户从暂存区中提取信息：以单选方式显示key，将暂存的数据显示在主界面的同时删除暂存区的当前记录

### 订货项：TODO

将本次订货信息写入订货周期表 Lieferant，以供计算应订情况

list1 显示应订计算结果：最后一次订货所在周 + 订货周期 < 本周

list2 显示当前未订货的门店或供货商，以供操作

list3 显示当前操作的门店或供货商，点击确认更改源文件

### 发票,账单,到货：

根据订货信息计算是否应进行操作

list1 显示计算结果：查找当前门店或供应商的订货列：已经标记

list2 显示当前项未标记的门店或供货商，以供操作

list3 显示当前操作的门店或供货商，点击确认更改源文件

### 入货,传真,投诉,原件：

根据到货信息计算是否应进行操作

list1 显示计算结果：查找当前门店或供应商的到货列：已经标记

list2 显示当前项未标记的门店或供货商，以供操作

list3 显示当前操作的门店或供货商，点击确认更改源文件

## TODO

以多线程辅助，提醒用户刷新数据，

点击查看开始计时，点击确认计时完毕


## __date__: 24.12.2021
## __Version__: 2.02 新增暂存区
未成功写入则弹窗提示暂存；成功写入则清空list3

数据暂存：储存操作、选项、列表等信息；允许调回

若关闭软件时存在暂存数据：弹窗提示，是否放弃暂存数据

## __date__: 23.12.2021
## __Version__: 2.01 更新FL、LF表，总览表中获取准确信息