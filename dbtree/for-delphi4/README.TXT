TxDBTree v1.1

Description:
  TxDBTree is derived from TTreeView, with capable reflect tree which stored
    in table.
  DBTreeView reading database table, type:
     ID  OwnerID  Text
     1   Null	      This is Root
       2	1	      This is Node1 on Root
       3	1	      This is Node2 on Root
       4	2	      This is Node3 on Node1

   and displaing the NEXT:

   This is Root
     | |
     | This is Node1 on Root
     |     |
     |     This is Node3 on Node1
     This is Node2 on Root

Contact Information:
  Feel free to contact me if you have any questions, comments or suggestions
  at buzaev@sbonline.ru.

Installation:
  To install, use "Component/Install Packages", press button "Add"
  and select file:
    in Delphi3 DBTree.dpl
    in Delphi4 DBTree4.bpl.

  И указать в Делфи путь к файлам xDBTree.pas и xDBTree.dcu

Note:
  Source is aviable on request.
  Price 15$.
