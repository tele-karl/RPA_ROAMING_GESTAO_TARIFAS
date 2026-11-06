import os
import subprocess
import tempfile
from src.services.log_service import registrar_log

def autenticar_rede_lagoa():
    r"""
    Autentica na rede \\lagoa.redecorp.br usando credenciais do .env e variável de ambiente.
    - Usuário: USUARIO_ICT (do .env)
    - Senha: RPA_ROAM_PASS (variável de ambiente Windows)
    
    Usa batch file com net use, passando senha via variável de ambiente.
    """
    try:
        usuario = os.getenv("USUARIO_ICT")
        senha = os.environ.get("RPA_ROAM_PASS")
        
        # Remover espaços em branco acidental (leading/trailing)
        if usuario:
            usuario = usuario.strip()
        if senha:
            senha = senha.strip()
        
        print(f"[DEBUG] Usuario: '{usuario}' (len={len(usuario) if usuario else 0})")
        print(f"[DEBUG] Senha: {'*' * len(senha) if senha else 'VAZIA'} (len={len(senha) if senha else 0})")
        
        if not usuario or not senha:
            registrar_log(
                "Iniciado", "Erro",
                f"Credenciais incompletas: USUARIO_ICT={bool(usuario)}, RPA_ROAM_PASS={bool(senha)}"
            )
            raise ValueError("Credenciais de autenticação de rede não encontradas")
        
        # Caminho UNC do servidor
        servidor_unc = "\\\\brtssisp003.redecorp.br\\DI"
        
        # Etapa 1: Desconectar TUDO (máxima limpeza)
        print("[INFO] Limpando todas as conexões de rede...")
        cmd_disconnect_all = 'net use * /delete /y'
        resultado_disconnect_all = subprocess.run(
            cmd_disconnect_all,
            shell=True,
            capture_output=True,
            text=True,
            timeout=10,
            stdin=subprocess.DEVNULL
        )
        print("[INFO] Todas as conexões removidas.")
        
        # Pequena pausa para garantir limpeza do cache
        import time
        time.sleep(2)
        
        # Etapa 2: Usar net use com senha via variável de ambiente
        print(f"[INFO] Conectando a {servidor_unc} com net use...")
        
        # Usar domínio REDECORP com o usuário
        usuario_dominio = f"REDECORP\\{usuario}"
        
        # Criar arquivo batch temporário que passa a senha via setlocal
        batch_script = f"""@echo off
setlocal enabledelayedexpansion
set "password=!RPA_TEMP_PASS!"
net use "{servidor_unc}" "!password!" /user:{usuario_dominio} /persistent:yes
exit /b %errorlevel%
"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.bat', delete=False) as f:
            f.write(batch_script)
            batch_path = f.name
        
        print(f"[DEBUG] Batch criado em: {batch_path}")
        print(f"[DEBUG] Usuario com domínio: {usuario_dominio}")
        
        try:
            # Preparar environment com senha temporária
            env = os.environ.copy()
            env['RPA_TEMP_PASS'] = senha
            
            # Executar batch
            print(f"[DEBUG] Executando batch: {batch_path}")
            resultado = subprocess.run(
                batch_path,
                capture_output=True,
                text=True,
                timeout=30,
                env=env,
                shell=True
            )
            
            print(f"[DEBUG] Retorno: {resultado.returncode}")
            if resultado.stdout:
                print(f"[DEBUG] stdout: {resultado.stdout}")
            if resultado.stderr:
                print(f"[DEBUG] stderr: {resultado.stderr}")
            
            if resultado.returncode == 0:
                registrar_log(
                    "Iniciado", "Sucesso",
                    f"Autenticação na rede lagoa realizada com sucesso para usuário: {usuario}"
                )
                print(f"[INFO] Conexão a {servidor_unc} estabelecida com sucesso.")
                return True
            else:
                erro = resultado.stderr.strip() if resultado.stderr else resultado.stdout.strip()
                registrar_log(
                    "Iniciado", "Erro",
                    f"Falha na autenticação de rede: {erro}"
                )
                print(f"[ERROR] Falha na conexão: {erro}")
                return False
                
        finally:
            # Limpar arquivo temporário
            try:
                os.unlink(batch_path)
            except:
                pass
            
    except subprocess.TimeoutExpired:
        registrar_log(
            "Iniciado", "Erro",
            "Timeout durante autenticação de rede - comando demorou muito"
        )
        print("[ERROR] Timeout durante autenticação.")
        return False
    except Exception as e:
        registrar_log(
            "Iniciado", "Erro",
            f"Exceção durante autenticação de rede: {str(e)}"
        )
        print(f"[ERROR] Exceção na autenticação: {e}")
        return False
