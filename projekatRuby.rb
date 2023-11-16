require "google_drive"

session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("12mxOw6rFaKvylKx4_3JHLRLO1tLnWgIQFJjfwLbCbK8").worksheets[0]
ws1 = session.spreadsheet_by_key("12mxOw6rFaKvylKx4_3JHLRLO1tLnWgIQFJjfwLbCbK8").worksheets[1]

class Tabela
  include Enumerable

  #pravimo konstruktor da cuvamo podatke iz nase tabele
  def initialize(worksheet)
    @worksheet = worksheet       #promenljiva za nas worksheet
    @tabela  = []            #nasa tabela
    @imenaKolona = []        #header, imena nasih kolona
    @indeksi = {}            #hash mapa de cuvamo sve nase kolone iz NasaKolona
    @pocetni_red = nil       # pocetni red tabele u worksheet posto tabela ne mora da se nadje u gornjem densom uglu
    @pocetna_kolona = nil    # pocetna kolona tabele u worksheet posto tabela ne mora da se nadje u gornjem densom uglu
    @prazni_redovi = []      # prazni redovi koji se nalaze izmedju headera i tabele ili redovi u kojima se nalazi subtotal cuvamo njihove indekse
    # da bi vratili nazad u tabelu
    @duzina_u_redovima
    @sirina_u_kolonama
  end

  attr_accessor :worksheet ,:tabela, :imenaKolona,:indeksi,:pocetna_kolona,:pocetni_red,:prazni_redovi,:duzina_u_redovima,:sirina_u_kolonama

  def kreirajTabelu
    @worksheet.num_rows.times do |red_indeks|
      @worksheet.num_cols.times do |kolona_indeks|
        vrednost_celije = @worksheet[red_indeks + 1,kolona_indeks + 1]

        unless vrednost_celije.nil? || vrednost_celije.empty?
          @pocetni_red = red_indeks + 1 if @pocetni_red.nil?
          @pocetna_kolona = kolona_indeks + 1 if @pocetna_kolona.nil?
          break
        end
      end
      break if @pocetni_red && @pocetna_kolona
    end
    @pocetni_red ||= 1
    @pocetna_kolona ||= 1
    @duzina_u_redovima = @worksheet.num_rows - @pocetni_red + 1
    @sirina_u_kolonama = @worksheet.num_cols - @pocetna_kolona + 1

  end

  def izvuciImenaKolonaIzWorksheeta(worksheet)
    @imena_kolona = []
    (1..worksheet.num_cols).each do |kolona_indeks|
      celija = worksheet[pocetni_red||1, kolona_indeks]
      if celija && !celija.empty?
        ime_kolone = celija&.strip&.split(' ')&.map.with_index { |rec, i| i.zero? ? rec.downcase : rec.capitalize }&.join
        @imena_kolona << ime_kolone unless ime_kolone.nil? || ime_kolone.empty?
      end
    end
    @imena_kolona
  end

  def uzmiIndekse
    return unless @imena_kolona && @pocetna_kolona && @sirina_u_kolonama
    (0..@sirina_u_kolonama - 1).each do |kolona|
      ime_kolone = @imena_kolona[kolona]
      indeks_kolone = @pocetna_kolona + kolona
      @indeksi[ime_kolone] = indeks_kolone
      puts "lalala #{@indeksi[ime_kolone]}"
    end
  end

  def sveSpojenoSveMalaSlova(string)
    string.downcase.split(' ').join('')
  end

  def popuniTabelu(ws)
    @tabela = []
    prazni_redovi = []
    ws.num_rows.times do |red_indeks|
      red = []
      ws.num_cols.times do |kolona_indeks|
        vrednost_celije = ws[red_indeks + 1, kolona_indeks + 1]
        unless vrednost_celije.nil? || vrednost_celije.empty?
          prazni_redovi << red_indeks + 1 and break if vrednost_celije.downcase.include?('total') || vrednost_celije.downcase.include?('subtotal')
          red << vrednost_celije unless vrednost_celije.nil? || vrednost_celije.empty? || prazni_redovi.include?(red_indeks + 1)
        end
      end
      next if red.empty? || prazni_redovi.include?(red_indeks + 1)

      @tabela << red
      #ne mogu da namestim za spojene
    end
    @tabela
  end

  def row(broj)
    return nil unless broj.is_a?(Integer) && @pocetni_red.is_a?(Integer)
    red_indeks = broj - @pocetni_red - 1
    return nil unless (0..@duzina_u_redovima - 1).cover?(red_indeks)

    @tabela[red_indeks]
    # @tabela[broj - @pocetni_red-1] if broj.between?(@pocetni_red, (@pocetni_red + @duzina_u_redovima) - 1)
  end

  def col(broj)
    return nil unless broj.between?(@pocetna_kolona, (@pocetna_kolona + @tabela[0].size - 1))
    kolona_indeks = broj - @pocetna_kolona
    @tabela.map { |red| red[kolona_indeks]  }
  end


  def each
    @tabela.each do |row|
      row.each do |cell|
        yield cell
      end
    end
  end

  def [](key)
    return nil if @indeksi.nil?

    column_index = @indeksi[key.to_s.downcase]
    return nil unless column_index

    column_data = col(column_index)
    return NasaKolona.new(self, column_data, column_index)
  end

  def []=(key, index, value)
    return if @indeksi.nil?
    column_index = @indeksi[key.to_s.downcase]
    return unless column_index

    row_index = index + @pocetni_red
    @tabela[row_index - @pocetni_red][column_index - @pocetna_kolona] = value
  end

  def +(druga_tabela)
    if (@tabela.row(4).eql?(druga_tabela.row(4)))
      @tabela.parse + druga_tabela.tabela.parse
    end
    return nil
  end

  def -(druga_tabela)
    if(@tabela.row(1).eql?(druga_tabela.table.row(1)))
      return @tabela.parse - druga_tabela.table.parse
    end
    return nil
  end


end



class NasaKolona
  def initialize(table, data, index)
    @table = table
    @data = data
    @index = index
  end

  attr_accessor :table,:index,:data

  def [](index)
    red_indeks = index + @table.pocetni_red
    puts"red indeks #{red_indeks}"
    return nil unless (1..@table.duzina_u_redovima).cover?(red_indeks)

    @table.row(red_indeks)[@index - @table.pocetna_kolona]
  end

  def []=(index, value)
    red_indeks = index + @table.pocetni_red
    return nil unless (1..@table.duzina_u_redovima).cover?(red_indeks)

    @table.tabela[red_indeks - @table.pocetni_red][@index - @table.pocetna_kolona] = value
  end




end
