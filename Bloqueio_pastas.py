import subprocess

pasta = r".\base_compilada"
usuario_permitido = "Zequinha"

# Remove herança
subprocess.run(
    f'icacls "{pasta}" /inheritance:r',
    shell=True
)

# Concede acesso total apenas ao usuário especificado
subprocess.run(
    f'icacls "{pasta}" /grant {usuario_permitido}:(OI)(CI)F',
    shell=True
)

print("Permissões aplicadas.")