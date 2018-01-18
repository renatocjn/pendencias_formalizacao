#/usr/bin/env ruby

require 'fox16'
require 'Clipboard'
require_relative "ProcessadorDePendencias"
include Fox

class ProcessadorDePendenciasGUI < FXMainWindow
  include ProcessadorDePendencias

  def initialize(app)
    # Invoke base class initialize first
    super(app, "Processador de pendencias", opts: DECOR_ALL)
	
    #Create a tooltip
    FXToolTip.new(self.getApp())
	
    #Controls on top
    controls = FXVerticalFrame.new(self, LAYOUT_SIDE_TOP|LAYOUT_FILL_X,
      padLeft: 40, padRight: 40, padTop: 10, padBottom: 10)
	  
    controlGroup = FXGroupBox.new(controls, "Selecione o banco sendo processado", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_SIDE_TOP|LAYOUT_FILL_X)
    optionsPopup = FXPopup.new(controlGroup)
    ["Detecção automática", "BMG", "Bradesco", "Santander"].each { |opt| FXOption.new(optionsPopup, opt) }
    @bank_select = FXOptionMenu.new controlGroup, optionsPopup, opts: FRAME_RAISED|FRAME_THICK|ICON_BEFORE_TEXT|LAYOUT_FILL_X
    
    controlGroup = FXGroupBox.new(controls, "Selecione a planilha a ser processada", GROUPBOX_TITLE_CENTER|FRAME_RIDGE|LAYOUT_SIDE_TOP|LAYOUT_FILL_X)
    FXButton.new(controlGroup, "Selecionar planilha", opts: LAYOUT_FILL_X|FRAME_THICK|FRAME_RAISED|BUTTON_DEFAULT).connect(SEL_COMMAND) do  
      dialog = FXFileDialog.new(self, "Selecione a planilha de pendencias")  
      dialog.selectMode = SELECTFILE_EXISTING
      dialog.patternList = ["Excel Files (*.xls,*.xlsx)"]  
      
      if dialog.execute != 0
        begin
          processedProposals = recoverProposalNumbersAndStateOfProposals(dialog.filename, @bank_select.current)
          insertProcessedProposalsToTable processedProposals
        rescue RuntimeError => err
          FXMessageBox::error self, MBOX_OK, "Algo deu errado...", err.message
        end
      end  
    end
    
    FXHorizontalSeparator.new(self,
        LAYOUT_SIDE_TOP|LAYOUT_FILL_X|SEPARATOR_GROOVE)
    
    # Status bar
    FXLabel.new(self, "contato: helpdesk@casebras.com.br", opts: JUSTIFY_RIGHT|LAYOUT_SIDE_BOTTOM)
    
    FXHorizontalSeparator.new(self,
        LAYOUT_SIDE_BOTTOM|LAYOUT_FILL_X|SEPARATOR_GROOVE)
    
    footer = FXVerticalFrame.new(self, LAYOUT_SIDE_BOTTOM|LAYOUT_FILL_X,
      padLeft: 40, padRight: 40, padTop: 10, padBottom: 10)
     
    FXButton.new(footer, "Copiar para área de transferência", opts: LAYOUT_SIDE_BOTTOM|FRAME_RAISED|FRAME_THICK|LAYOUT_FILL_X).connect(SEL_COMMAND) do 
      xmlFriendlyContent = String.new
      @table.numRows.times { |r| xmlFriendlyContent += "#{@table.getItemText(r,0)}\t#{@table.getItemText(r,1)}\n" }
      Clipboard.copy xmlFriendlyContent
    end
    
    # Contents
    contents = FXHorizontalFrame.new(self,
      LAYOUT_CENTER_X|FRAME_NONE|LAYOUT_FILL_X|LAYOUT_FILL_Y, :padding => 10)
     
    @table = FXTable.new(contents, opts: LAYOUT_CENTER_X|LAYOUT_CENTER_Y|LAYOUT_FILL_X|LAYOUT_FILL_Y|TABLE_ROW_RENUMBER)
    
    #@table.borderColor = FXRGB(255, 255, 255)
    @table.visibleRows = 5
    @table.visibleColumns = 2
    @table.setTableSize(10, 2)
    @table.editable = false
    @table.setColumnText(0, "Propostas")
    @table.setColumnText(1, "UF")
    @table.setItemJustify(0, 0, FXTableItem::CENTER_X|FXTableItem::CENTER_Y)
    @table.setItemJustify(0, 1, FXTableItem::CENTER_X|FXTableItem::CENTER_Y)
  end
  
  def create
    super
    show(PLACEMENT_SCREEN)
  end
  
  def insertProcessedProposalsToTable (processedProposals)
    @table.setTableSize(0,2)
    @table.setColumnText(0, "Propostas")
    @table.setColumnText(1, "UF")
    processedProposals.each do |proposal, uf|
      @table.insertRows(0)
      @table.setItemText(0,0, Integer(proposal).to_s)
      @table.setItemText(0,1, uf.to_s)
	end
  end
end

app = FXApp.new "ProcessadorDePendencias", "Processamento de planilhas de pendencias"
ProcessadorDePendenciasGUI.new app
app.create
app.run