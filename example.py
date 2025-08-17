import agate

import agateods

table = agate.Table.from_ods('examples/test.ods')

print(table)
table.print_table()
