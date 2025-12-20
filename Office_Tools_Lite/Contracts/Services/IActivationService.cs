namespace Office_Tools_Lite.Contracts.Services;

public interface IActivationService
{
    Task ActivateAsync(object activationArgs);
}
