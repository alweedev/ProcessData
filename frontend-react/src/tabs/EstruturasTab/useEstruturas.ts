import { useMemo, useState } from "react";
import { pushToast } from "../../toast/toastStore";

export interface PreviewItem {
  aprovacaoId: string;
  aprovacaoPor: string;
  aprovacao?: string;
  tipo?: string;
  valor?: string;
  ccCodigo?: string;
  ccDescricao?: string;
  viajanteNomeCompleto?: string;
  posicoes?: string[];
  segundoNivel?: boolean;
  ficaraSemAprovador?: boolean;
}

interface Approver {
  nomeCompleto?: string;
  cpf?: string;
}

interface Summary {
  estruturasAfetadas?: number;
  ocorrenciasTotal?: number;
  porAprovacaoPor?: Record<string, number>;
}

interface EmptyStructure {
  aprovacaoId: string;
}

function normalizeCpf(value: string): string {
  return (value || "").replace(/\D/g, "");
}

export function useEstruturas() {
  const [usersFile, setUsersFile] = useState<File | null>(null);
  const [baseFile, setBaseFile] = useState<File | null>(null);
  const [cpf, setCpf] = useState("");
  const [removeSecondLevel, setRemoveSecondLevel] = useState(false);

  const [status, setStatus] = useState<{ message: string; isError: boolean }>({ message: "", isError: false });
  const [approver, setApprover] = useState<Approver | null>(null);
  const [summary, setSummary] = useState<Summary | null>(null);
  const [items, setItems] = useState<PreviewItem[]>([]);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());

  const [previewLoading, setPreviewLoading] = useState(false);
  const [exportLoading, setExportLoading] = useState(false);

  const hasPreview = summary !== null;

  const readinessHint = useMemo(() => {
    if (!usersFile && !baseFile) return "Envie a base de usuários e a base de aprovação para começar.";
    if (usersFile && !baseFile) return "Base de usuários selecionada. Agora selecione a base de aprovação.";
    if (!usersFile && baseFile) return "Base de aprovação selecionada. Agora selecione a base de usuários.";
    return "Bases prontas. Informe o CPF e clique em Verificar.";
  }, [usersFile, baseFile]);

  function toggleSelected(id: string, checked: boolean) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (checked) next.add(id);
      else next.delete(id);
      return next;
    });
  }

  function toggleSelectAll(checked: boolean) {
    setSelectedIds(checked ? new Set(items.map((i) => i.aprovacaoId)) : new Set());
  }

  function clearAll() {
    setUsersFile(null);
    setBaseFile(null);
    setCpf("");
    setRemoveSecondLevel(false);
    setApprover(null);
    setSummary(null);
    setItems([]);
    setSelectedIds(new Set());
    setStatus({ message: "", isError: false });
    pushToast("Todos os dados foram limpos.", "info");
  }

  async function preview() {
    if (!usersFile || !baseFile) {
      setStatus({ message: "Envie a base de usuários e a base de aprovação.", isError: true });
      pushToast("Envie a base de usuários e a base de aprovação.", "danger");
      return;
    }
    const cpfDigits = normalizeCpf(cpf);
    if (cpfDigits.length !== 11) {
      setStatus({ message: "CPF inválido. Informe 11 dígitos.", isError: true });
      pushToast("CPF inválido. Informe 11 dígitos.", "danger");
      return;
    }

    setPreviewLoading(true);
    setStatus({ message: "Verificando estruturas...", isError: false });
    try {
      const fd = new FormData();
      fd.append("users_file", usersFile);
      fd.append("base_file", baseFile);
      fd.append("cpf", cpf);
      const res = await fetch("/api/aprovacao/remover/preview", { method: "POST", body: fd });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const msg = data?.error || "Erro ao obter preview.";
        setStatus({ message: msg, isError: true });
        pushToast(msg, "danger");
        setItems([]);
        setSummary(null);
        return;
      }
      const newItems: PreviewItem[] = data.items || [];
      setItems(newItems);
      setSelectedIds(new Set(newItems.map((i) => i.aprovacaoId)));
      setApprover(data.approver || null);
      setSummary(data.summary || {});
      if (!newItems.length) {
        setStatus({ message: "Nenhuma estrutura encontrada para o CPF informado.", isError: true });
        pushToast("Nenhuma estrutura encontrada para o CPF informado.", "info");
      } else {
        setStatus({ message: `Encontradas ${newItems.length} estrutura(s) contendo o aprovador.`, isError: false });
      }
    } catch {
      const msg = "Erro ao contatar o servidor de aprovação.";
      setStatus({ message: msg, isError: true });
      pushToast(msg, "danger");
    } finally {
      setPreviewLoading(false);
    }
  }

  async function doExport(mode: "all" | "selected", ignoreEmptyWarning = false): Promise<EmptyStructure[] | null> {
    if (!usersFile || !baseFile) {
      setStatus({ message: "Envie novamente a base de usuários e a base de aprovação.", isError: true });
      pushToast("Envie novamente a base de usuários e a base de aprovação.", "danger");
      return null;
    }
    const cpfDigits = normalizeCpf(cpf);
    if (cpfDigits.length !== 11) {
      setStatus({ message: "CPF inválido. Informe 11 dígitos.", isError: true });
      pushToast("CPF inválido. Informe 11 dígitos.", "danger");
      return null;
    }
    if (mode === "selected" && selectedIds.size === 0) {
      setStatus({ message: "Nenhuma estrutura selecionada. Marque pelo menos uma linha.", isError: true });
      pushToast("Nenhuma estrutura selecionada.", "danger");
      return null;
    }

    setExportLoading(true);
    setStatus({ message: "Gerando base de aprovação atualizada...", isError: false });
    try {
      const fd = new FormData();
      fd.append("users_file", usersFile);
      fd.append("base_file", baseFile);
      fd.append("cpf", cpf);
      fd.append("mode", mode);
      if (removeSecondLevel) fd.append("remove_second_level", "1");
      if (ignoreEmptyWarning) fd.append("ignore_empty_warning", "1");
      if (mode === "selected") {
        for (const id of selectedIds) fd.append("selected_aprovacao_ids[]", id);
      }

      const res = await fetch("/api/aprovacao/remover/export", { method: "POST", body: fd });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        if (errData.warning && errData.estruturasSemAprovador?.length) {
          return errData.estruturasSemAprovador as EmptyStructure[];
        }
        const msg = errData?.error || "Erro ao gerar base de aprovação atualizada.";
        setStatus({ message: msg, isError: true });
        pushToast(msg, "danger");
        return null;
      }

      const blob = await res.blob();
      if (!blob || !blob.size) {
        const msg = "Arquivo retornado vazio ou inválido.";
        setStatus({ message: msg, isError: true });
        pushToast(msg, "danger");
        return null;
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "base_aprovacao_atualizada.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setStatus({ message: "Base de aprovação atualizada gerada com sucesso.", isError: false });
      pushToast("Base de aprovação atualizada gerada com sucesso.", "success");
      return null;
    } catch {
      const msg = "Erro ao baixar a base de aprovação atualizada.";
      setStatus({ message: msg, isError: true });
      pushToast(msg, "danger");
      return null;
    } finally {
      setExportLoading(false);
    }
  }

  return {
    usersFile,
    setUsersFile,
    baseFile,
    setBaseFile,
    cpf,
    setCpf,
    removeSecondLevel,
    setRemoveSecondLevel,
    status,
    approver,
    summary,
    items,
    selectedIds,
    toggleSelected,
    toggleSelectAll,
    hasPreview,
    readinessHint,
    previewLoading,
    exportLoading,
    preview,
    doExport,
    clearAll,
  };
}
