import genanki
import re
import math

class Cloze:
  with open('assets/style.css', encoding='utf-8') as f:
    cssLoaded = f.read()
    f.close()
    # print(cssLoaded)
  with open('assets/front.html', encoding='utf-8') as f:
    frontLoaded = f.read()
    f.close()
    # print(cssLoaded)
  with open('assets/back.html', encoding='utf-8') as f:
    backLoaded = f.read()
    f.close()
    # print(cssLoaded)
  my_model = genanki.Model(
    112252932,
    'Cloze (drquochoai excel to anki)',
    model_type=genanki.Model.CLOZE,
    fields=[
      {'name': 'id'},
      {'name': 'Front'},
      { 'name': "Extra note"},
      { 'name': "Title"},
      { 'name': "Footer"}
    ],
    templates=[
      {
        'name': 'Cloze',
        'qfmt': frontLoaded,
        'afmt': backLoaded,
      },
    ],
    css= cssLoaded
  )
  my_deck = []
  def resetDeck(self):
    self.my_deck.clear()
  def createDeck(self, deckName, mota):
    if mota == '':
      mota = "Tạo bởi Trần Quốc Hoài with love"
    self.my_deck.append(
      genanki.Deck(
      deck_id= int( str(int.from_bytes(deckName.encode(), 'little'))[:18] ),
      name=deckName,
      description=mota)
      )
    # Chuyển name thành interger
    # https://stackoverflow.com/questions/31701991/string-of-text-to-unique-integer-method
  def addNote(self, fields, myguids):
    my_note = genanki.Note(
      model= self.my_model,
      fields = fields, guid= myguids)
    self.my_deck[-1].add_note(my_note)
    
  def saveAnkiPackage(self, packageName):
    genanki.Package(self.my_deck).write_to_file(packageName + '.apkg')
  def htmlProcess(self, text_cell, text_cell_xf, text_cell_runlist, font_list, col_idx):    
    htmlValue = ''
    countCloze = 1
    # print(text_cell_runlist)
    if text_cell_runlist:
        # print ('(cell multi style) SEGMENTS:')
        segments = []
        for segment_idx in range(len(text_cell_runlist)):
            start = text_cell_runlist[segment_idx][0]
            # the last segment starts at given 'start' and ends at the end of the string
            end = None
            if segment_idx != len(text_cell_runlist) - 1:
                end = text_cell_runlist[segment_idx + 1][0]
            segment_text = text_cell[start:end]
            """ segments.append({
                'text': segment_text,
                'font': book.font_list[text_cell_runlist[segment_idx][1]]
            }) """
            # xử lý theo font chữ
            font = font_list[text_cell_runlist[segment_idx][1]]
            tempHtml = segment_text
            if font.bold == 1 :
                # https://xlrd.readthedocs.io/en/latest/api.html#xlrd.formatting.Font
                if col_idx == 0:
                  tempHtml = '<span class="dr_bold">{{c' + str(countCloze) + '::' + tempHtml + '}}</span>'
                  countCloze += 1
                else:
                  tempHtml = '<span class="dr_bold">' + tempHtml + '</span>'
            if font.italic == 1:
                tempHtml = '<span class="dr_italic">' + tempHtml + '</span>'
            if font.underlined == 1:
                tempHtml = '<span class="dr_underlined">' + tempHtml + '</span>'
            htmlValue += tempHtml
        # segments did not start at beginning, assume cell starts with text styled as the cell
        if text_cell_runlist[0][0] != 0:
            """ segments.insert(0, {
                'text': text_cell[:text_cell_runlist[0][0]],
                'font': book.font_list[text_cell_xf.font_index]
            }) """
            # Thằng đầu tiên sẽ không bắt đầu từ số 0
            htmlValue = text_cell[:text_cell_runlist[0][0]] + htmlValue
    else:
        """ print ('(cell single style)')
        print( 'italic:', book.font_list[text_cell_xf.font_index].italic,)
        print ('bold:', book.font_list[text_cell_xf.font_index].bold) """
        htmlValue = text_cell
        font = font_list[text_cell_xf.font_index]
        if font.bold == 1 :
            # https://xlrd.readthedocs.io/en/latest/api.html#xlrd.formatting.Font
            htmlValue = '<span class="dr_bold">' + htmlValue + '</span>'
        if font.italic == 1:
            htmlValue = '<span class="dr_italic">' + htmlValue + '</span>'
        if font.underlined == 1:
            htmlValue = '<span class="dr_underlined">' + htmlValue + '</span>'

    #  tìm tất cả link http từ htmlValue và thay thế bằng img
    # https://stackoverflow.com/questions/9760588/how-do-you-extract-a-url-from-a-string-using-python
    # the_list = re.compile(r'hello(?i)') ; bỏ qua case ví dụ Http
    htmlValue = htmlValue.replace("\\n", "<br/>")
    allHttps = re.findall(r'(https?:\/\/.*\.(?:png|jpg|jpeg|gif|png|svg))(?i)', htmlValue)
    if len(allHttps) > 0:
      # print(allHttps)
      for link in allHttps:
        htmlValue = htmlValue.replace(link, '<img class="dr_img" src="' + link + '"/>')
    # Nếu không có cloze thì thêm cloze ở cuối html
    if col_idx == 0 and htmlValue.find("\:\:") == -1:
      htmlValue += '<span class="hide">{{c1::autocloze}}</span>'
    
    return htmlValue
