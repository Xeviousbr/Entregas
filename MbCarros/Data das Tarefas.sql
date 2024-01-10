Alter Table Tarefas Add Column DtConclusao Date
Alter Table TarefasTemp Add Column DtConclusao Date
Update Tarefas Set DtConclusao = Pago Where pago is not null
Update Tarefas Set DtConclusao = Now Where pago is null and situacao = 3