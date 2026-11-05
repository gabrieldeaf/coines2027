# Gera o teste.html a partir do index.html: python3 ferramentas/gerar_teste.py index.html ferramentas/teste_extra.html teste.html
import sys
src, extra, dst = sys.argv[1:4]
html = open(src, encoding='utf-8').read()
bloco = open(extra, encoding='utf-8').read()
style, resto = bloco.split('</style>', 1)
barra, script = resto.split('<script>', 1)
def troca(h, velho, novo):
    assert h.count(velho) == 1, velho
    return h.replace(velho, novo)
html = troca(html, '<title>Inscrição COINES 2027</title>', '<title>Teste visual COINES 2027</title>\n<meta name="robots" content="noindex">')
html = troca(html, '</style>', style.replace('<style>', '') + '</style>')
html = troca(html, '<body>', '<body>\n' + barra.strip())
html = troca(html, '<script>\n// ======== CONFIGURAÇÃO', '<script>' + script.replace('</script>', '').rstrip() + '\n</script>\n<script>\n// ======== CONFIGURAÇÃO')
# Login do teste separado do login real
html = troca(html, "const CHAVE = 'coines27_codigo';", "const CHAVE = 'coines27_teste';")
open(dst, 'w', encoding='utf-8').write(html)
print('ok', dst)
