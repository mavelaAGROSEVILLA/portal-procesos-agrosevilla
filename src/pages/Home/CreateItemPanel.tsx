import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import type { IDepartment, IProcessGroup } from '../../models/ProcessModels';
import { ProcessGroupService } from '../../services/ProcessGroupService';
import { ProcessService } from '../../services/ProcessService';
import styles from '../../webparts/processView/components/Home.module.scss';

export type CreateItemMode = 'group' | 'process';

interface ICreateItemPanelProps {
  mode: CreateItemMode;
  sp: SPFI;
  departments: IDepartment[];
  groups: IProcessGroup[];
  selectedDepartmentId?: number;
  selectedGroupId?: number;
  onClose: () => void;
  onCreated: () => Promise<void>;
}

const CreateItemPanel: React.FunctionComponent<ICreateItemPanelProps> = ({
  mode,
  sp,
  departments,
  groups,
  selectedDepartmentId,
  selectedGroupId,
  onClose,
  onCreated
}) => {
  const [code, setCode] = React.useState('');
  const [name, setName] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [departmentId, setDepartmentId] = React.useState<number | undefined>(selectedDepartmentId);
  const [groupId, setGroupId] = React.useState<number | undefined>(selectedGroupId);
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string>();

  const departmentOptions: IDropdownOption[] = departments.map(department => ({ key: department.id, text: `${department.code} - ${department.name}` }));
  const departmentGroups = groups.filter(group => group.departmentId === departmentId);
  const groupOptions: IDropdownOption[] = departmentGroups.map(group => ({ key: group.id, text: `${group.code} - ${group.name}` }));

  const submit = async (): Promise<void> => {
    if (!code.trim() || !name.trim() || !departmentId) {
      setError(mode === 'group'
        ? 'Completa el código, el nombre y el departamento.'
        : 'Completa el código, el nombre y el departamento.');
      return;
    }
    try {
      setSaving(true);
      setError(undefined);
      if (mode === 'group') {
        await new ProcessGroupService(sp).createGroup(code.trim(), name.trim(), departmentId);
      } else {
        const selectedDepartment = departments.find(department => department.id === departmentId);
        await new ProcessService(sp).createProcess(
          code.trim(),
          name.trim(),
          description.trim(),
          '',
          groupId as number,
          departmentId,
          selectedDepartment ? selectedDepartment.code : ''
        );
      }
      await onCreated();
    } catch (saveError) {
      console.error('[ProcessView] Error al crear elemento', saveError);
      const errorMessage = saveError instanceof Error ? saveError.message : 'Error desconocido';
      setError(`No se ha podido guardar: ${errorMessage}`);
    } finally {
      setSaving(false);
    }
  };

  return (
    <section className={styles.createPanel} aria-labelledby="create-panel-heading">
      <div className={styles.detailHeader}>
        <div>
          <p className={styles.eyebrow}>Nuevo registro</p>
          <h3 id="create-panel-heading">{mode === 'group' ? 'Añadir grupo de procesos' : 'Añadir proceso'}</h3>
        </div>
        <DefaultButton text="Cancelar" onClick={onClose} disabled={saving} />
      </div>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div className={styles.formGrid}>
        <TextField label="Código" required value={code} onChange={(_, value) => setCode(value || '')} disabled={saving} />
        <TextField label={mode === 'group' ? 'Nombre del grupo' : 'Nombre del proceso'} required value={name} onChange={(_, value) => setName(value || '')} disabled={saving} />
        <Dropdown label="Departamento" required options={departmentOptions} selectedKey={departmentId} onChange={(_, option) => {
          setDepartmentId(option?.key as number);
          setGroupId(undefined);
        }} disabled={saving} />
        {mode === 'process' && <>
          <Dropdown label="Grupo de procesos (opcional)" options={groupOptions} selectedKey={groupId} onChange={(_, option) => {
            const selectedGroup = departmentGroups.find(group => group.id === option?.key);
            setGroupId(selectedGroup?.id);
            setDepartmentId(selectedGroup?.departmentId || departmentId);
          }} disabled={saving || departmentId === undefined || departmentGroups.length === 0} placeholder={departmentId === undefined ? 'Selecciona primero un departamento' : 'Selecciona un grupo'} />
          <TextField className={styles.formWide} label="Descripción" multiline rows={4} value={description} onChange={(_, value) => setDescription(value || '')} disabled={saving} />
        </>}
      </div>
      <div className={styles.formActions}>
        <PrimaryButton text={saving ? 'Guardando...' : 'Guardar'} onClick={() => { submit().catch(() => undefined); }} disabled={saving} />
        {saving && <Spinner size={SpinnerSize.small} ariaLabel="Guardando" />}
      </div>
    </section>
  );
};

export default CreateItemPanel;