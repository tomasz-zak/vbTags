# Create a tags file for VBA programs, usable with vi.
# Tagged are:
# - Subprocedures
# - Functions (even inside other defs or classes)
# Warns about files it cannot open.
# No warnings about duplicate tags.

import sys, re, os, pdb

tags = []    # Modified global variable!

expr = '[^E]*(Sub|Function) (\w+)'
matcher = re.compile(expr)

def main():
    args = sys.argv[1:]
    for filename in args:
        treat_file(filename)
    if tags:
        fp = open('tags', 'w')
        tags.sort()
        for s in tags: fp.write(s)

def treat_file(filename):
    try:
        fp = open(filename, 'r')
    except:
        sys.stderr.write('Cannot open %s\n' % filename)
        return
    base = os.path.basename(filename)
    if base[-4:] == '.bas':
        base = base[:-4]
    s = base + '\t' + filename + '\t' + '1\n'
    tags.append(s)
    while 1:
        line = fp.readline()
        if not line:
            break
        m = matcher.match(line)
        if m:
	    # pdb.set_trace()
            content = m.group(0)
            name = m.group(2)
            s = name + '\t' + filename + '\t/^' + content + '/\n'
            tags.append(s)

if __name__ == '__main__':
    main()
