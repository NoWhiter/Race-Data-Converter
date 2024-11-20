 # GBC及osu!比赛成绩整理工具

这是一个适用于 **GBC** 和其他 osu! 比赛中导出的 **CSV 文件**进行整理与排序的离线应用程序。

---

## **功能简介**
- **支持去重操作**：可选是否对选手的成绩进行去重。
- **多表格导出**：自动生成包含选手成绩、图排名以及团队总分等信息的多张表格。
- **支持团队赛**：可额外添加队伍信息，按队伍整理比赛数据。

---

## **使用方法**
1. **选择是否去重**  
   根据需要，勾选是否对选手成绩进行去重操作。

2. **导入 CSV 文件**  
   选择导出的比赛数据文件（格式为 `.csv`）。

3. **设置 Excel 导出路径**  
   指定整理后的数据保存路径（格式为 `.xlsx`）。

4. **团队赛设置（可选）**  
   如果是团队赛，可以添加队伍信息，格式如下：
     ```
     选手A, 队伍名
     选手B, 队伍名
     ……
     ```

5. **点击开始**  
   程序将生成三或四张 Excel 表格，其中：
   - **选手成绩总览**：展示每位选手的整体成绩情况。
   - **单图排名**：详细展示每张图的个人排名。
   - **单图分数**：详细展示每张图的个人分数。
   - **团队赛总分（若适用）**：统计团队赛每个队伍的单图总分。

---

## **注意事项**
- **关闭相关文件**  
  请确保在运行程序前，已关闭所有正在使用的 CSV 文件和 Excel 文件。

- **运行环境**  
  需安装.NET 8.0及以上版本。
=======
# GBC鍙妎su!姣旇禌鎴愮哗鏁寸悊宸ュ叿

杩欐槸涓�涓�傜敤浜?**GBC** 鍜屽叾浠?osu! 姣旇禌涓鍑虹殑 **CSV 鏂囦欢**杩涜鏁寸悊涓庢帓搴忕殑绂荤嚎搴旂敤绋嬪簭銆?

---

## **鍔熻兘绠�浠?*
- **鏀寔鍘婚噸鎿嶄綔**锛氬彲閫夋槸鍚﹀閫夋墜鐨勬垚缁╄繘琛屽幓閲嶃�?
- **澶氳〃鏍煎鍑?*锛氳嚜鍔ㄧ敓鎴愬寘鍚�夋墜鎴愮哗銆佸浘鎺掑悕浠ュ強鍥㈤槦鎬诲垎绛変俊鎭殑澶氬紶琛ㄦ牸銆?
- **鏀寔鍥㈤槦璧?*锛氬彲棰濆娣诲姞闃熶紞淇℃伅锛屾寜闃熶紞鏁寸悊姣旇禌鏁版嵁銆?

---

## **浣跨敤鏂规硶**
1. **閫夋嫨鏄惁鍘婚噸**  
   鏍规嵁闇�瑕侊紝鍕鹃�夋槸鍚﹀閫夋墜鎴愮哗杩涜鍘婚噸鎿嶄綔銆?

2. **瀵煎叆 CSV 鏂囦欢**  
   閫夋嫨瀵煎嚭鐨勬瘮璧涙暟鎹枃浠讹紙鏍煎紡涓?`.csv`锛夈�?

3. **璁剧疆 Excel 瀵煎嚭璺緞**  
   鎸囧畾鏁寸悊鍚庣殑鏁版嵁淇濆瓨璺緞锛堟牸寮忎负 `.xlsx`锛夈�?

4. **鍥㈤槦璧涜缃紙鍙�夛級**  
   濡傛灉鏄洟闃熻禌锛屽彲浠ユ坊鍔犻槦浼嶄俊鎭紝鏍煎紡濡備笅锛?
     ```
     閫夋墜A, 闃熶紞鍚?
     閫夋墜B, 闃熶紞鍚?
     鈥︹�?
     ```

5. **鐐瑰嚮寮�濮?*  
   绋嬪簭灏嗙敓鎴愪笁鎴栧洓寮?Excel 琛ㄦ牸锛屽寘鎷細
   - **閫夋墜鎴愮哗鎬昏**锛氬睍绀烘瘡浣嶉�夋墜鐨勬暣浣撴垚缁╂儏鍐点�?
   - **鍗曞浘鎺掑悕**锛氳缁嗗睍绀烘瘡寮犲浘鐨勪釜浜烘帓鍚嶃�?
   - **鍗曞浘鍒嗘暟**锛氳缁嗗睍绀烘瘡寮犲浘鐨勪釜浜哄垎鏁般�?
   - **鍥㈤槦璧涙�诲垎锛堣嫢閫傜敤锛?*锛氱粺璁″洟闃熻禌姣忎釜闃熶紞鐨勫崟鍥炬�诲垎銆?

---

## **娉ㄦ剰浜嬮」**
- **鍏抽棴鐩稿叧鏂囦欢**  
  璇风‘淇濆湪杩愯绋嬪簭鍓嶏紝宸插叧闂墍鏈夋鍦ㄤ娇鐢ㄧ殑 CSV 鏂囦欢鍜?Excel 鏂囦欢銆?

- **杩愯鐜**  
  闇�瀹夎.NET 8.0鍙婁互涓婄増鏈�?
>>>>>>> 96cc6b42769f486300e9b0e99fe6d65bcae1e891

---
