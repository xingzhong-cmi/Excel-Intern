/**
 * Excel助手 - Frontend Application (Vue 3)
 */
const { createApp, ref } = Vue;

createApp({
    setup() {
        // State
        const currentStep = ref(1);
        const isDragging = ref(false);
        const uploading = ref(false);
        const uploadedFile = ref(null);
        const instruction = ref('');
        const processing = ref(false);
        const processingPhase = ref(1);
        const result = ref(null);
        const showScript = ref(false);
        const errorMessage = ref('');
        const fileInput = ref(null);

        // --- Upload Methods ---
        function triggerFileInput() {
            fileInput.value?.click();
        }

        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (file) uploadFile(file);
        }

        function handleDrop(event) {
            isDragging.value = false;
            const file = event.dataTransfer.files[0];
            if (file) uploadFile(file);
        }

        async function uploadFile(file) {
            // Validate file type
            const validExtensions = ['.xlsx', '.xls', '.csv'];
            const ext = '.' + file.name.split('.').pop().toLowerCase();
            if (!validExtensions.includes(ext)) {
                showError('不支持的文件格式，请上传 .xlsx、.xls 或 .csv 文件');
                return;
            }

            // Validate file size (50MB)
            if (file.size > 50 * 1024 * 1024) {
                showError('文件大小超过50MB限制');
                return;
            }

            uploading.value = true;
            uploadedFile.value = null;

            const formData = new FormData();
            formData.append('file', file);

            try {
                const response = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData,
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || '上传失败');
                }

                const data = await response.json();
                uploadedFile.value = data;
            } catch (err) {
                showError(err.message || '文件上传失败，请重试');
            } finally {
                uploading.value = false;
                // Reset file input
                if (fileInput.value) fileInput.value.value = '';
            }
        }

        function removeFile() {
            uploadedFile.value = null;
        }

        // --- Instruction Methods ---
        function useExample(text) {
            instruction.value = text;
        }

        async function processInstruction() {
            if (!uploadedFile.value || !instruction.value.trim()) return;

            processing.value = true;
            processingPhase.value = 1;
            result.value = null;

            // Simulate processing phases for UX
            const phaseTimer1 = setTimeout(() => { processingPhase.value = 2; }, 2000);
            const phaseTimer2 = setTimeout(() => { processingPhase.value = 3; }, 5000);

            try {
                const response = await fetch('/api/process', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        filenames: [uploadedFile.value.filename],
                        instruction: instruction.value.trim(),
                    }),
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || '处理请求失败');
                }

                const data = await response.json();
                result.value = data;
                currentStep.value = 3;
            } catch (err) {
                result.value = {
                    success: false,
                    message: err.message || '处理失败，请重试',
                    script: '',
                    output_files: [],
                    preview: null,
                };
                currentStep.value = 3;
            } finally {
                clearTimeout(phaseTimer1);
                clearTimeout(phaseTimer2);
                processing.value = false;
            }
        }

        // --- Navigation ---
        function goToStep(step) {
            if (step === 2 && !uploadedFile.value) {
                showError('请先上传文件');
                return;
            }
            currentStep.value = step;
        }

        function startOver() {
            currentStep.value = 1;
            uploadedFile.value = null;
            instruction.value = '';
            result.value = null;
            showScript.value = false;
        }

        // --- Utility ---
        function showError(message) {
            errorMessage.value = message;
            setTimeout(() => { errorMessage.value = ''; }, 4000);
        }

        return {
            currentStep,
            isDragging,
            uploading,
            uploadedFile,
            instruction,
            processing,
            processingPhase,
            result,
            showScript,
            errorMessage,
            fileInput,
            triggerFileInput,
            handleFileSelect,
            handleDrop,
            removeFile,
            useExample,
            processInstruction,
            goToStep,
            startOver,
        };
    },
}).mount('#app');
