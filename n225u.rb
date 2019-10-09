#!/usr/bin/ruby -Ku
# -*- coding: utf-8 -*-
# Yahooファイナンスから225銘柄の終値をダウンロード
# 0.1			2009/6/4	muz
# 0.1a Yafooファイナンスの仕様が変わって入力したコードの順で値が出てこないようになったので、ソートする          2009/9/25 
# 0.1b 9205 OUT, 9022 IN.       2010/1/22
# 0.1c 銘柄入替 5407IN          2010/3/29
# 0.1d 銘柄入替 5020,8630       2010/4/2
# 0.1e 銘柄入替 3404out,5214in  2010/9/28
# 0.1f 銘柄入替 6796out,8804in  2010/10/01
# 0.1g 銘柄入替 6764,6991,8403out, 6506,7735,8750in  2011/03/29
# 0.1h 8404,8606 OUT, 8304,8729 IN.       2011/8/29
# 0.1i 9737 OUT, 6113 IN.       2011/9/28
# 0.2  値段と差を出力する.      2012/4/13
# 0.2a Yahooファイナンスのコマンドが変わってしまったので、それに対応した。
#								2012/8/15
# 0.2b 実行に失敗するとSTDERRに"Failed"を出力するようにした。2012/8/16
# 0.2c 株価のデータが225でなく200になってしまうことがあった。2012/8/20
# 0.2d コマンドラインの'&'がエスケープしてなかったため、cygwin環境で実行するとオプションが伝わらず誤動作していた。 2012/09/27
# 0.2d 銘柄入れ替え. 5413, 5703 IN  2012/10/02
# 0.2e log機能追加 2012/12/27
# 0.2f Core2Duoへの引っ越しに伴いPATHを変更　2013/01/17
# 0.2g 銘柄入れ替え. 3893 OUT 2013/03/27
# 0.2h 銘柄入れ替え. 3863 IN  2013/04/02
# 0.2i 銘柄入れ替え. 6988 IN, 3289 OUT, みなし額面変更 8銘柄 2013/09/26
# 0.2j 銘柄入れ替え. 3289 IN, 3864 OUT 2013/10/02
# 0.2k 銘柄入れ替え. 1334 OUT 4503,4543 みなし額面変更 2014/03/28
# 0.2l 銘柄入れ替え. 1333 IN  2014/04/02
# 0.3  TOPIX100銘柄の小数点化により小数点以下1桁を出力。 2014/07/22
# 0.31a DOS窓でもCygwinでも同じ動きをするようにした。
#      デバッグレベルを加えクラスの実装にした。 2014/07/24
# 0.31b Yahoo側の仕様変更?で一度に50こ取れないようになったので、 step = 30にした。stepが変わってもループ側が変更してなかったので修正した。2014/07/31
# 0.32 またもやYahoo financeの仕様変更に伴い、改造。2015/01/08
# 0.33 timeout実装. 2015/0804
# 0.33b 銘柄入れ替え. 3110,8803 OUT, 1808,2432 IN  2015/10/01
# 0.34 utf8に変更. cygwinでもDOS窓でも動くようにした。 2015/12/16
# 0.35 指定した証券コードをtxtから探すようにした。
#      class CodeValueにした。
#      証券コードが存在しないとき、ログに記録しabort.       2016/03/30
# 0.35a 7186IN 2016/04/04
# 0.36 コードと価格をスキャンするところの変更
# 0.36b 6753OUT,7272IN 2016/08/01
# 0.36c 4041OUT,4755IN 2016/10/03
# 0.36d 6767OUT,4578IN 2017/01/24
# 0.36e 6502OUT,6724IN 2017/08/01
# 0.36f 3865,6508OUT,6098,6178IN 2017/10/02
# 0.36g 5413,4631IN 2018/12/26
# 0.36h 6773OUT,6645IN 2019/03/18
# 0.36i 5002OUT,5019 2019/03/27
# 0.36j 6366OUT,7832 2019/08/01
# 0.36k 9681OUT,2413IN 2019/10/01
Version="0.36k"

# #5413(日新製鋼ホールディングス)の値が返ってこない事がある。？2012/12/12
# コード順にソートして結果を出力するのでコードの順は任意。ワークシートのコードは昇順であること。
# DOS窓からはwindows,cygtermからはcygwinのrubyが起動される。
# windowsから駆動するときは、'cmd.exe xxxx.rb'
# コマンドの例http://info.finance.yahoo.co.jp/search/?query=1332+1334+1605+1721+1801&d=v1&k=c3&h=on&z=m&esearch=1
# コマンドが変更になったらブラウザーで設定し、コマンドを調べる。
require 'stringio'
require 'timeout'
#====================================================================
class Log
  def initialize(name)
    @logname = name
  end

  def print(s ="")
    STDERR.puts s
    logfp = File.open(@logname, "a")
    t = Time.now.strftime("[%Y/%m/%d %H:%M:%S]") + s
    logfp.puts t
    logfp.close
  end
end
#====================================================================
class CodeValue
  def initialize
  end

  def findCodeVal(code, txt)
=begin 大きな文字列txtからcodeを探し、その値を文字列で返す。文字列はstringioで開く。

例: code="8316"
  • (株)三井住友フィナンシャルグループ [8316] - 東証1部掲示板

    15:00
  •

    3,471
        -91（-2.55%）
        時価総額4,908,187百万円
戻り値:3471
=end

    sio = StringIO.new(txt, "r")

    puts "code=#{code}" if $debug>=2
    puts "name=#{$1}" if $debug>=2
    puts "txt=#{txt}" if $debug>=3
 
    r1 = Regexp.new(/\s+(.*)\s*\[#{code}\] - 東証1部/)
    while l=sio.gets do
        break if l =~ r1  # 目的の証券コードまで読み飛ばす.
    end
    puts "name=#{$1}" if $debug>=2

    if l==nil then  # not found
      $log.print("code=#{code} Not found")
      abort("txt is as follows:
===============================================
#{txt}
===============================================")
    end

    r2 = Regexp.new(/^\s*([\d,\.]+)$/)
    l=sio.gets.chomp while l !~ r2
    if not l then
      $log.print("price Not found")
      $log.printabort("price Not found")
    end
    price = $1.sub(/,/,"")

    l=sio.gets.chomp                 # 前日との差は次の行にあるはず
    if l =~ /((-|\+)?[\d.]+)（.*）/  # 例     +2（+1.55%）
      val = $1
      if val=~ /\./ then dif = val.to_f
      else
        dif=val.to_i
      end
    elsif l =~  /---（0.00%/ then    # 前日との差が0
      dif = 0
    else
      $log.print("dif not found")    # 不明
      exit
    end


    STDERR.puts "#{code} #{price} #{dif}"
    sio.close
    price
  end
#------------------------------------------------
  def getval(code)
    v = Hash.new
    url = $base0 + code.join('.t+')+'.t' + $base1
    command = $w3m + ' ' + url
    $log.print("command=#{command}") if $debug >= 1

    txt = ""
    begin 
      timeout($command_timeout) {
        txt = `#{command}` # command内の文字列がここで評価されるので、'&'などはエスケープしてあること。 2012/09/27
      }
    rescue => exe
      $log.print("ERROR:Timeout:#{exe}")
      $log.print(txt)
      abort("ERROR:Timeout:#{exe}")
    end
    $log.print(txt) if $debug>=3

    code.each {|c|
      v.merge!({c=>findCodeVal(c,txt)})
    }
    v
  end
end
#====================================================================
#============== BEGIN_OF_MAIN =======================================
$debug = 0
progname=$0
while ARGV.length > 0  do
  x = ARGV.shift
  if x =~ /-d/ or x =~ /--debug/ then
    abort("-d need #") if ARGV.length == 0 or (x=ARGV.shift) !~ /\d+/
    $debug = x.to_i
  elsif x =~ /-h/ or x =~ /--help/
    STDERR.puts "n225.rb [-d|--debug #][-h|--help]
-d|--debug #:debug level # 0=none 1=std 2=verbose
-h|--help  :help "
    exit
  end
end
#---------------------- BEGIN_OF_CONSTANTS -----------------------
#MY_LIB_PATH="C:/zzz/"		 この形式または
#MY_LIB_PATH="C:\\zzz\\"   この形式ならばWindowsもmocygwinも両方いける。
#が、コマンドを実行するときのPATHはwindowsとcygwinで違うので、両方を分ける。
# ENV['windir']はwindowsだと"C:\Windows" cygwinだと""  ... のはずだったが、同じ。ENV['LANG']で分ける.

if ENV['LANG'] == "ja_JP.UTF-8" then				        # cygwin,Linux
  $MY_LIB_PATH='/home/muz/'
  $w3m = '/usr/bin/w3m -dump '
#  $w3m = '/usr/bin/w3m -dump -cols 240'
  $outfile = '/home/muz/zzz225.csv'
  logname = '/home/muz/n225.log'
else       # dos窓
  $MY_LIB_PATH="C:\\zzz\\"						# windows流
  #MY_LIB_PATH='C:\zzz\'						# だめ。なぜ？ 2013/01/17
  $w3m = "C:\\cygwin\\bin\\w3m.exe -dump -cols 240"
  $outfile = "C:\\zzz\\zzz225.csv"
  logname = "C:\\zzz\\n225.log"
end

$log = Log.new(logname)
$log.print("#{progname} #{Version} start")
#$log.print("ENV['windir']=#{ENV['windir']}")
#
# 2013/03/27 3893 out
# 2013/04/02 3863 in 
#if false
if true
code = [ \
5020,8630,  #2010/4/2kara
1808,2432,
1332,1333,1605,1721,1801,1802,1803,1812,1925,1928,1963,2002,2269,2282,2413,2501,2502,2503,2531,2768,2801,2802,2871,2914,3086,3099,3101,3103,3105,3289,3382,3401,3402,3405,3407,3436,3861,3863,4004,4005,4021,4042,4043,4061,4063,4151,4183,4188, \

4208,4272,4324,4452,4502,4503,4506,4507,4519,4523,4543,4568,4578,4631,4689,4704,4751,4755,4901,4902,4911,5019,5101,5108,5201,5202,5214,5232,5233,5301,5332,5333,5401,5406,5411,5541,5631,5703,5706,5707,5711,5713,5714,5801,5802,5803,5901,6098,6103,6113,6178,6301,6302,6305, \

6326,6361,6367,6471,6472,6473,6479,6501,6503,6504,6506,6645,6674,6701,6702,6703,6724,6752,6758,6762,6770,6841,6857,6902,6952,6954,6971,6976,6988,7003,7004,7011,7012,7013,7186,7201,7202,7203,7205,7211,7261,7267,7269,7270,7272,7731,7733, \

7735,7751,7752,7762,7832,7911,7912,7951,8001,8002,8015,8028,8031,8035,8053,8058,8233,8252,8253,8267,8303,8303,8304,8306,8308,8309,8316,8331,8354,8355,8411,8601,8604,8628,8725,8729,8750,8766,8795,8801,8802,8804,8830,9001,9005,9007,9008,9009,9020, \

9021,9062,9064,9101,9104,9107,9202,9022,9301,9412,9432,9433,9437,9501,9502,9503,9531,9532,9602,9613,9735,9766,9983,9984 ]
else
code = [1332,1333,1605,1721,1801]
end

# base0 + "code1 + code2+ ..." + base1 をgetする.
$base0 = 'http://info.finance.yahoo.co.jp/search/?query='
$base1 = '\&ei=UTF-8\&view=1'		# '&'をエスケープする。

step = 13		# 15 ずつ
$command_timeout=10

#---------------------- END_OF_CONSTANTS -----------------------
begin
  aFile = File.new($outfile, 'w')
rescue => ex
  $log.print(ex.message)
  abort ex.message
end

value = Hash.new
cv = CodeValue.new
result="Failed:"		# 実行結果（暫定版)
0.upto(((225+0.0)/step).ceil){ | i |
  cd = code.slice(i*step, step)
  break unless cd
  value.merge!(cv.getval(cd))
}

$log.print("Size = #{value.size}") if $debug >= 1
if value.size != 225 then
  $log.print("size not equal 225") 
end

value.keys.sort.each { |k|
  aFile.printf("%0.1f\n", value[k])  # TOPIX100銘柄の小数点化により
  $log.print("#{k}: #{value[k]}") if $debug >= 2
}

$log.print("#{progname} #{Version} end")
exit(0)
#================ END_OF_MAIN =======================================
