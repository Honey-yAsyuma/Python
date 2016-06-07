import argparse
import MeCab
import xlsxwriter


parser = argparse.ArgumentParser()
parser.add_argument("txtname")

args = parser.parse_args()


try:
    f = open(args.txtname)

except IOError as e:
	print("ファイルが存在しません")
	exit()

else:
	import traceback
	traceback.print_exc()
	exit()


mecab = MeCab.Tagger('-Ochasen')
data = f.read()
node = mecab.parseToNode(data)
phrases = node.next


dict = {}
while phrases:
	try:
		k = node.surface
		node = node.next

		if dict.get(k) is not None:
			dict[k] = dict[k] + 1
		else:
			dict[k] = 1

	except AttributeError:
		break


workbook = xlsxwriter.Workbook("Mecab_result.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write("A1","コーパス")
worksheet.write("B1","出現回数")


i = 1
for k,v in sorted( dict.items(), key=lambda x:x[1],reverse = True):
	print ("word:" + k, " count:" + str(v))
	i += 1
	w_place_A = "A{0}".format(str(i))
	w_place_B = "B{0}".format(str(i))
	worksheet.write(w_place_A,k)
	worksheet.write(w_place_B,v)

f.close()
workbook.close()
