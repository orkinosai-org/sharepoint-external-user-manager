import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Modal,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  CommandBar,
  ICommandBarItemProps,
  Dropdown,
  IDropdownOption,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter
} from '@fluentui/react';
import { IExternalLibrary, IExternalUser } from '../models/IExternalLibrary';

export interface IManageUsersModalProps {
  isOpen: boolean;
  library: IExternalLibrary | null;
  onClose: () => void;
  onAddUser: (libraryId: string, email: string, permission: 'Read' | 'Contribute' | 'Full Control') => Promise<void>;
  onRemoveUser: (libraryId: string, userId: string) => Promise<void>;
  onGetUsers: (libraryId: string) => Promise<IExternalUser[]>;
  onSearchUsers: (query: string) => Promise<IExternalUser[]>;
}

export interface IAddUserFormData {
  email: string;
  permission: 'Read' | 'Contribute' | 'Full Control';
}

export const ManageUsersModal: React.FC<IManageUsersModalProps> = ({
  isOpen,
  library,
  onClose,
  onAddUser,
  onRemoveUser,
  onGetUsers,
  onSearchUsers
}) => {
  const [users, setUsers] = useState<IExternalUser[]>([]);
  const [selectedUsers, setSelectedUsers] = useState<IExternalUser[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [operationMessage, setOperationMessage] = useState<{ message: string; type: MessageBarType } | null>(null);
  
  // Add User Form State
  const [showAddUserForm, setShowAddUserForm] = useState<boolean>(false);
  const [addUserForm, setAddUserForm] = useState<IAddUserFormData>({
    email: '',
    permission: 'Read'
  });
  const [addingUser, setAddingUser] = useState<boolean>(false);
  const [validationErrors, setValidationErrors] = useState<{[key: string]: string}>({});
  
  // Remove User Confirmation
  const [showRemoveConfirmation, setShowRemoveConfirmation] = useState<boolean>(false);
  const [removingUser, setRemovingUser] = useState<boolean>(false);

  const [selection] = useState(new Selection({
    onSelectionChanged: () => {
      setSelectedUsers(selection.getSelection() as IExternalUser[]);
    }
  }));

  // Load users when modal opens
  useEffect(() => {
    if (isOpen && library) {
      loadUsers();
    } else if (!isOpen) {
      // Reset state when modal closes
      setUsers([]);
      setSelectedUsers([]);
      setError('');
      setOperationMessage(null);
      setShowAddUserForm(false);
      setShowRemoveConfirmation(false);
      selection.setAllSelected(false);
    }
  }, [isOpen, library]);

  const loadUsers = async (): Promise<void> => {
    if (!library) return;
    
    setLoading(true);
    setError('');
    setOperationMessage(null);
    
    try {
      const loadedUsers = await onGetUsers(library.id);
      setUsers(loadedUsers);
      
      if (loadedUsers.length === 0) {
        setOperationMessage({
          message: 'No external users found for this library.',
          type: MessageBarType.info
        });
      }
    } catch (err) {
      const errorMessage = err.message || 'Failed to load users';
      setError(errorMessage);
    } finally {
      setLoading(false);
    }
  };

  const handleAddUser = async (): Promise<void> => {
    if (!library) return;
    
    if (!validateAddUserForm()) {
      return;
    }

    setAddingUser(true);
    setError('');
    
    try {
      await onAddUser(library.id, addUserForm.email.trim(), addUserForm.permission);
      
      setOperationMessage({
        message: `Successfully added ${addUserForm.email} to ${library.name}`,
        type: MessageBarType.success
      });
      
      // Reset form and reload users
      setAddUserForm({ email: '', permission: 'Read' });
      setShowAddUserForm(false);
      await loadUsers();
      
    } catch (err) {
      setError(err.message || 'Failed to add user');
    } finally {
      setAddingUser(false);
    }
  };

  const handleRemoveUsers = async (): Promise<void> => {
    if (!library || selectedUsers.length === 0) return;
    
    setRemovingUser(true);
    setError('');
    
    try {
      // Remove each selected user
      for (const user of selectedUsers) {
        await onRemoveUser(library.id, user.id);
      }
      
      setOperationMessage({
        message: `Successfully removed ${selectedUsers.length} user(s) from ${library.name}`,
        type: MessageBarType.success
      });
      
      setShowRemoveConfirmation(false);
      setSelectedUsers([]);
      selection.setAllSelected(false);
      await loadUsers();
      
    } catch (err) {
      setError(err.message || 'Failed to remove users');
    } finally {
      setRemovingUser(false);
    }
  };

  const validateAddUserForm = (): boolean => {
    const errors: {[key: string]: string} = {};
    
    // Email validation
    if (!addUserForm.email.trim()) {
      errors.email = 'Email address is required';
    } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(addUserForm.email.trim())) {
      errors.email = 'Please enter a valid email address';
    }
    
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleInputChange = (field: keyof IAddUserFormData) => 
    (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
      setAddUserForm(prev => ({
        ...prev,
        [field]: newValue || ''
      }));
      
      // Clear validation error for this field
      if (validationErrors[field]) {
        setValidationErrors(prev => {
          const newErrors = { ...prev };
          delete newErrors[field];
          return newErrors;
        });
      }
    };

  const handlePermissionChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      setAddUserForm(prev => ({
        ...prev,
        permission: option.key as 'Read' | 'Contribute' | 'Full Control'
      }));
    }
  };

  const handleClose = (): void => {
    if (!addingUser && !removingUser && !loading) {
      onClose();
    }
  };

  // Define columns for the users list
  const columns: IColumn[] = [
    {
      key: 'displayName',
      name: 'Name',
      fieldName: 'displayName',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IExternalUser) => (
        <Text variant="medium">{item.displayName || 'Unknown'}</Text>
      )
    },
    {
      key: 'email',
      name: 'Email',
      fieldName: 'email',
      minWidth: 200,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IExternalUser) => (
        <Text variant="small">{item.email}</Text>
      )
    },
    {
      key: 'permissions',
      name: 'Permissions',
      fieldName: 'permissions',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IExternalUser) => (
        <Text variant="small">{item.permissions}</Text>
      )
    },
    {
      key: 'invitedDate',
      name: 'Invited Date',
      fieldName: 'invitedDate',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IExternalUser) => (
        <Text variant="small">
          {item.invitedDate.toLocaleDateString()}
        </Text>
      )
    }
  ];

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'addUser',
      text: 'Add User',
      iconProps: { iconName: 'AddFriend' },
      onClick: () => setShowAddUserForm(true)
    },
    {
      key: 'removeUser',
      text: 'Remove User',
      iconProps: { iconName: 'RemoveFriend' },
      disabled: selectedUsers.length === 0,
      onClick: () => setShowRemoveConfirmation(true)
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: loadUsers
    }
  ];

  const permissionOptions: IDropdownOption[] = [
    { key: 'Read', text: 'Read' },
    { key: 'Contribute', text: 'Contribute' },
    { key: 'Full Control', text: 'Full Control' }
  ];

  const modalProps = {
    isOpen,
    onDismiss: handleClose,
    isBlocking: addingUser || removingUser || loading,
    containerClassName: 'manage-users-modal'
  };

  return (
    <>
      <Modal {...modalProps}>
        <div style={{ padding: '20px', minWidth: '600px', maxWidth: '800px' }}>
          <Stack tokens={{ childrenGap: 20 }}>
            <Stack.Item>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack>
                  <Text variant="xLarge" styles={{ root: { fontWeight: 'semibold' } }}>
                    Manage External Users
                  </Text>
                  <Text variant="medium" styles={{ root: { color: '#666', marginTop: '4px' } }}>
                    {library ? `Library: ${library.name}` : 'No library selected'}
                  </Text>
                </Stack>
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  ariaLabel="Close"
                  onClick={handleClose}
                  disabled={addingUser || removingUser || loading}
                />
              </Stack>
            </Stack.Item>

            {error && (
              <Stack.Item>
                <MessageBar messageBarType={MessageBarType.error}>
                  {error}
                </MessageBar>
              </Stack.Item>
            )}

            {operationMessage && (
              <Stack.Item>
                <MessageBar messageBarType={operationMessage.type}>
                  {operationMessage.message}
                </MessageBar>
              </Stack.Item>
            )}

            <Stack.Item>
              <CommandBar
                items={commandBarItems}
                ariaLabel="User Management Actions"
              />
            </Stack.Item>

            <Stack.Item>
              {loading ? (
                <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                  <Spinner size={SpinnerSize.large} label="Loading users..." />
                </Stack>
              ) : (
                <DetailsList
                  items={users}
                  columns={columns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selection={selection}
                  selectionPreservedOnEmptyClick={true}
                  selectionMode={SelectionMode.multiple}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="select row"
                />
              )}
            </Stack.Item>

            <Stack.Item>
              <Stack horizontal tokens={{ childrenGap: 15 }}>
                <Text variant="small" styles={{ root: { color: '#666' } }}>
                  Total External Users: {users.length}
                </Text>
                <Text variant="small" styles={{ root: { color: '#666' } }}>
                  Selected: {selectedUsers.length}
                </Text>
              </Stack>
            </Stack.Item>

            {/* Add User Form */}
            {showAddUserForm && (
              <Stack.Item>
                <div style={{ 
                  border: '1px solid #edebe9', 
                  borderRadius: '4px', 
                  padding: '16px',
                  backgroundColor: '#faf9f8'
                }}>
                  <Stack tokens={{ childrenGap: 15 }}>
                    <Stack.Item>
                      <Text variant="mediumPlus" styles={{ root: { fontWeight: 'semibold' } }}>
                        Add External User
                      </Text>
                    </Stack.Item>
                    
                    <Stack.Item>
                      <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <Stack.Item grow>
                          <TextField
                            label="Email Address *"
                            value={addUserForm.email}
                            onChange={handleInputChange('email')}
                            disabled={addingUser}
                            errorMessage={validationErrors.email}
                            placeholder="Enter user's email address"
                            description="Enter the email address of the external user to invite"
                          />
                        </Stack.Item>
                        <Stack.Item>
                          <Dropdown
                            label="Permission Level *"
                            options={permissionOptions}
                            selectedKey={addUserForm.permission}
                            onChange={handlePermissionChange}
                            disabled={addingUser}
                            styles={{ dropdown: { width: 150 } }}
                          />
                        </Stack.Item>
                      </Stack>
                    </Stack.Item>

                    <Stack.Item>
                      <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                          text={addingUser ? 'Adding...' : 'Add User'}
                          onClick={handleAddUser}
                          disabled={addingUser || !addUserForm.email.trim()}
                          iconProps={addingUser ? undefined : { iconName: 'AddFriend' }}
                        />
                        <DefaultButton
                          text="Cancel"
                          onClick={() => {
                            setShowAddUserForm(false);
                            setAddUserForm({ email: '', permission: 'Read' });
                            setValidationErrors({});
                          }}
                          disabled={addingUser}
                        />
                      </Stack>
                    </Stack.Item>
                  </Stack>
                </div>
              </Stack.Item>
            )}

            <Stack.Item>
              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <DefaultButton
                  text="Close"
                  onClick={handleClose}
                  disabled={addingUser || removingUser || loading}
                />
              </Stack>
            </Stack.Item>
          </Stack>
        </div>
      </Modal>

      {/* Remove User Confirmation Dialog */}
      <Dialog
        hidden={!showRemoveConfirmation}
        onDismiss={() => setShowRemoveConfirmation(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Remove External Users',
          subText: `Are you sure you want to remove ${selectedUsers.length} user(s) from "${library?.name}"? This action cannot be undone.`
        }}
        modalProps={{
          isBlocking: removingUser
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleRemoveUsers}
            text={removingUser ? 'Removing...' : 'Remove'}
            disabled={removingUser}
          />
          <DefaultButton
            onClick={() => setShowRemoveConfirmation(false)}
            text="Cancel"
            disabled={removingUser}
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};