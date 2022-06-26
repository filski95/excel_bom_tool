# excel_bom_tool
python tool counting consumption of materials for different anchors.

-----
production was provided with an excel extract from the german system with specific anchors. Sales team was meant to receive quick feedback if it is possible to produce given amount within a certain time frame. 

Due to the system constraints each anchor had to be checked manually (bill of material) before production could be 100% sure that materials are available.
In was very frequent to have about 20-30 different anchors in each quote, manual checkes were very time consuming and prone to errors.
The amount of material depended on lengths of certain parts of the anchors and number of strands in them. 

![image](https://user-images.githubusercontent.com/93283105/175822051-449a71f8-e0dc-45e5-932e-34f895e1b829.png)


The tool opens up an excel, clears it up -> to make sure data is consistent with the script and subsequently reads off product description and gets specific lengths with a use of RE module.

Description:
Litzentemporäranker 0,6"
48 Stk. à 19,5 m, Lv 6m
Lfr 12,5m, Üli 1m, Üre 0m![image](https://user-images.githubusercontent.com/93283105/175822326-bc12fe19-c9fc-4703-be5d-ba829ba760de.png)


After that, materials are being counted based on the BOM dictionary in the tool.

