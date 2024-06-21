# -*- coding: utf-8 -*-
""" Renumber Mark Parameter of MEP elements."""

__title__ = 'MarkMaster'
__author__ = 'André Rodrigues da Silva'
		
from pyrevit import revit, DB, forms

def get_elements_by_ambiente(doc, ambiente_value):
    # Coletar todos os elementos no documento
    collector = DB.FilteredElementCollector(doc)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    # Filtrar elementos pelo parâmetro "AMBIENTE"
    matching_elements = []
    for element in elements:
        param = element.LookupParameter("AMBIENTE")
        if param and param.AsString() == ambiente_value:
            matching_elements.append(element)
    
    return matching_elements

def renumber_elements(elements):
    for index, element in enumerate(elements, start=1):
        param = element.LookupParameter("Mark")
        if param:
            param.Set(str(index))

def create_schedule(doc, category, ambiente_value):
    # Iniciar transação
    t = DB.Transaction(doc, "Create Schedule")
    t.Start()
    
    # Criar tabela
    schedule = DB.ViewSchedule.CreateSchedule(doc, category.Id)
    
    # Adicionar parâmetros à tabela
    fields = schedule.Definition.GetSchedulableFields()
    
    def add_field(field_name):
        for field in fields:
            if field.GetName(doc) == field_name:
                schedule.Definition.AddField(field)
                break
    
    # Adicionar campos desejados
    add_field("Mark")
    add_field("Type")
    add_field("Comments")
    
    # Filtrar por ambiente
    field = schedule.Definition.GetField(DB.ScheduleFieldType.Instance, DB.ElementId(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS))
    filter = DB.ScheduleFilter(field.FieldId, DB.ScheduleFilterType.Equal, ambiente_value)
    schedule.Definition.AddFilter(filter)
    
    # Finalizar transação
    t.Commit()
    
    return schedule

# Obter documento do Revit
doc = revit.doc

# Solicitar ao usuário o valor do parâmetro AMBIENTE
ambiente_value = forms.ask_for_string("Enter the 'AMBIENTE' value:")

if not ambiente_value:
    forms.alert("No 'AMBIENTE' value provided!", exitscript=True)

# Obter elementos pelo valor do parâmetro AMBIENTE
elements = get_elements_by_ambiente(doc, ambiente_value)

if not elements:
    forms.alert(f"No elements found with 'AMBIENTE' value '{ambiente_value}'", exitscript=True)

# Numeração dos elementos
t = DB.Transaction(doc, "Renumber Elements")
t.Start()
renumber_elements(elements)
t.Commit()

# Obter a categoria dos elementos (assumindo que todos são da mesma categoria)
category = elements[0].Category

# Criar tabela (schedule)
schedule = create_schedule(doc, category, ambiente_value)

# Exibir mensagem de confirmação
forms.alert(f"Elements renumbered and schedule '{schedule.Name}' created successfully!", exitscript=True)
